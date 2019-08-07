Imports MySql.Data.MySqlClient

Public Class ADA_Set

    'This class represents an entire ADA set which is equivalent to ADAXXX in the ADA spreadsheet (setup tab)

    Private ADA_set_name As String
    Public Motors_table(47, 3) As String
    Public inputs_table(18, 3) As String
    Public push_table(18, 3) As String
    Public lights_table(14, 3) As String
    Public field_table(11, 3) As String
    Public bus_table(11, 3) As String
    Public scanner_table(3, 3) As String
    Public brakes_table(14, 3) As String

    'My global settings
    Public Voltage As String
    Public efficiency As String
    Public Power_factor As String
    Public include_ADA_brackets As String
    Public Use_4pts_IO As String
    Public Use_8pts_IO As String
    Public Use_16pts_IO As String
    Public removal_msbox As String
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


    '---------------- DEVICE TYPE INPUTS -----------------
    Public Panel_bool As Integer
    ' Public R2_setup As Double  'Power distribution
    Public NEMA12 As Double
    Public NEMA4X As Double
    Public PLC_GateWay_PGW As Double
    Public ControlLogix_PLC_Box As Double
    Public CompactLogix_PLC_Box As Double
    Public PLC_UPS_Battery_Backup As Double
    Public Amp_Mon_Electro As Double
    Public Remote_IO As Double
    Public Amp_30_Drops As Double
    Public MDRs As Double
    Public Total_current_per_group As Double

    '---------------- DISCONNECTS INPUTS -----------------
    Public ADA_NEMA12_BOX24_No_Starters As Double
    Public ADA_NEMA12_BOX30_No_Starters As Double
    Public motor_disconnect As Double


    '---------- amps used (STARTER SIZE SETUP)-------------------------

    Public Results_amps_motors(11, 15) As Double  'store amp result for motors (Reversing and non reversing)
    Public Results_amps_vfd(11, 15) As Double  'store amps result for vfd

    Public voltage_index As Integer ' 0 if 230, 1 if 480 and 2 if 575


    '---------- List of ADA Parts qtys (This list will store the qtys calculate of each ADA part. They are dictionary list so we can access the qty by ADA part name

    Public Starter_Panel As New Dictionary(Of String, Double)
    Public IO_Panel As New Dictionary(Of String, Double)
    Public PLC_Panel As New Dictionary(Of String, Double)
    Public Field As New Dictionary(Of String, Double)
    Public Scanners As New Dictionary(Of String, Double)
    '  Public M12 As New Dictionary(Of String, Double)
    '  Public M12_ES As New Dictionary(Of String, Double)
    ' Public M12_SWD As New Dictionary(Of String, Double)

    '----------------------------------------
    Public non_reversing_starters As Double
    Public reversing_starters As Double
    Public VFD_starters As Double
    Public Soft_starters As Double
    Public Starters As Double
    Public big_panels As Double
    Public PLC As Double
    Public busbar_slots As Double
    Public non_reversing_amps_p As Double
    Public reversing_amps_p As Double
    Public vfd_amps_p As Double
    Public soft_amps_p As Double



    '----------------///////////////// VARIABLES USED IN STARTER PANEL TAB /////////////----------------------

    Public total_current As Double   'total current (H318)
    Public Primary_PS_5A As Double
    Public Primary_PS_10A As Double
    Public Primary_PS_20A As Double
    Public Primary_PS_40A As Double
    Public Panel_Space_Top As Double
    Public Panel_Space_Bottom As Double
    Public DC_amps_main_PS As Double
    Public DC_amps_aux_PS As Double
    Public Aux_PS_5A As Double
    Public Aux_PS_10A As Double
    Public Aux_PS_20A As Double
    Public Aux_PS_40A As Double
    Public Total_swire_nodes As Double

    '--------------------------------------------------------------------------------------------------------------------

    '---------------////////////////////// VARIABLES USED IN IO PANEL TAB ///////////////////////------------------

    Public Inputs As Double
    Public Monitoring_Inputs As Double
    Public Non_Motion_Outputs As Double
    Public motion_outputs As Double
    Public has_wire As Double
    Public system_power_req As Double

    '---- the variables below are used in CR assigmnet-----

    Public cr_input As Double
    Public cr_n_motion As Double
    Public cr_motion As Double

    '-------------I/O 4 Point Cards

    'IN
    Public IO_4_inputs As Double
    Public IO_4_non_motion As Double
    Public IO_4_motion As Double

    Public Totals_IO_4 As Double
    Public IO_4_convert_4 As Double
    Public IO_4_convert_8 As Double
    Public IO_4_ADA_SWD_FF_Coupling As Double
    Public IO_4_Pre_combine_8s As Double
    Public IO_4_total_JB As Double

    Public IO_4_pigtail_single As Double
    Public IO_4_conn_At_JB_single As Double
    Public IO_4_pigtail_double As Double
    Public IO_4_conn_At_JB_double As Double

    'OUT
    Public Totals_IO_4_out As Double
    Public IO_4_convert_4_out As Double
    Public IO_4_convert_8_out As Double
    Public IO_4_convert_16_out As Double
    Public IO_4_Pre_combine_8s_out As Double
    Public IO_4_total_JB_out As Double

    Public IO_4_pigtail_single_out As Double
    Public IO_4_conn_At_JB_single_out As Double
    Public IO_4_pigtail_double_out As Double
    Public IO_4_conn_At_JB_double_out As Double


    '------------------I/O 8 Points Cards

    'IN
    Public IO_8_inputs As Double
    Public IO_8_non_motion As Double
    Public IO_8_motion As Double

    Public Totals_IO_8 As Double
    Public IO_8_convert_4 As Double
    Public IO_8_convert_8 As Double
    Public IO_8_convert_16 As Double
    Public IO_8_Pre_combine_8s As Double
    Public IO_8_total_JB As Double

    Public IO_8_pigtail_single As Double
    Public IO_8_conn_At_JB_single As Double
    Public IO_8_pigtail_double As Double
    Public IO_8_conn_At_JB_double As Double

    'OUT
    Public Totals_IO_8_out As Double
    Public IO_8_convert_4_out As Double
    Public IO_8_convert_8_out As Double
    Public IO_8_convert_16_out As Double
    Public IO_8_Pre_combine_8s_out As Double
    Public IO_8_total_JB_out As Double

    Public IO_8_pigtail_single_out As Double
    Public IO_8_conn_At_JB_single_out As Double
    Public IO_8_pigtail_double_out As Double
    Public IO_8_conn_At_JB_double_out As Double

    '--------------------- I/O 16 Point Cards

    'IN
    Public IO_16_inputs As Double
    Public IO_16_non_motion As Double
    Public IO_16_motion As Double

    Public Totals_IO_16 As Double
    Public IO_16_convert_4 As Double
    Public IO_16_convert_8 As Double
    Public IO_16_convert_16 As Double
    Public IO_16_Pre_combine_8s As Double
    Public IO_16_total_JB As Double

    Public IO_16_pigtail_single As Double
    Public IO_16_conn_At_JB_single As Double
    Public IO_16_pigtail_double As Double
    Public IO_16_conn_At_JB_double As Double

    'OUT
    Public Totals_IO_16_out As Double
    Public IO_16_convert_4_out As Double
    Public IO_16_convert_8_out As Double
    Public IO_16_convert_16_out As Double
    Public IO_16_Pre_combine_8s_out As Double
    Public IO_16_total_JB_out As Double

    Public IO_16_pigtail_single_out As Double
    Public IO_16_conn_At_JB_single_out As Double
    Public IO_16_pigtail_double_out As Double
    Public IO_16_conn_At_JB_double_out As Double

    '------------------------------------------------------------------------------

    Public Sub New(set_name As String, motor_data As Array, inputs_data As Array, push_data As Array, lights_data As Array, field_data As Array, bus_data As Array, scanner_data As Array, brakes_data As Array)  'Constructor

        Me.ADA_set_name = set_name
        Motors_table = motor_data.Clone
        inputs_table = inputs_data.Clone
        push_table = push_data.Clone
        lights_table = lights_data.Clone
        field_table = field_data.Clone
        bus_table = bus_data.Clone
        scanner_table = scanner_data.Clone
        brakes_table = brakes_data.Clone

        '--------------- Boolean values in Device Type specifications -----------------------

        'If Not ADA_Setup.Panel_bool_drop.SelectedItem Is Nothing Then
        '    'Panel bool
        '    If String.Equals(ADA_Setup.Panel_bool_drop.SelectedItem.ToString, "Yes") Then
        '        Panel_bool = 1
        '    Else
        '        Panel_bool = 0
        '    End If
        'Else
        '    Panel_bool = 0
        'End If


        'If Not ADA_Setup.Nema12_com.SelectedItem Is Nothing Then
        '    'NEMA 12
        '    If String.Equals(ADA_Setup.Nema12_com.SelectedItem.ToString, "Yes") Then
        '        NEMA12 = 1
        '    Else
        '        NEMA12 = 0
        '    End If
        'Else
        '    NEMA12 = 0
        'End If

        'If Not ADA_Setup.nema4x_box.SelectedItem Is Nothing Then
        '    'NEMA 4x
        '    If String.Equals(ADA_Setup.nema4x_box.SelectedItem.ToString, "Yes") Then
        '        NEMA4X = 1
        '    Else
        '        NEMA4X = 0
        '    End If
        'Else
        '    NEMA4X = 0
        'End If

        'If Not ADA_Setup.PLC_gate_box.SelectedItem Is Nothing Then
        '    'PLC Gateway
        '    If String.Equals(ADA_Setup.PLC_gate_box.SelectedItem.ToString, "Yes") Then
        '        PLC_GateWay_PGW = 1
        '    Else
        '        PLC_GateWay_PGW = 0
        '    End If
        'Else
        '    PLC_GateWay_PGW = 0
        'End If

        'If Not ADA_Setup.control_box.SelectedItem Is Nothing Then
        '    'Control logic
        '    If String.Equals(ADA_Setup.control_box.SelectedItem.ToString, "Yes") Then
        '        ControlLogix_PLC_Box = 1
        '    Else
        '        ControlLogix_PLC_Box = 0
        '    End If
        'Else
        '    ControlLogix_PLC_Box = 0
        'End If

        'If Not ADA_Setup.compact_box.SelectedItem Is Nothing Then
        '    'Compact logic
        '    If String.Equals(ADA_Setup.compact_box.SelectedItem.ToString, "Yes") Then
        '        CompactLogix_PLC_Box = 1
        '    Else
        '        CompactLogix_PLC_Box = 0
        '    End If
        'Else
        '    CompactLogix_PLC_Box = 0
        'End If

        'If Not ADA_Setup.PLC_UPS_box.SelectedItem Is Nothing Then
        '    'PLC UPS
        '    If String.Equals(ADA_Setup.PLC_UPS_box.SelectedItem.ToString, "Yes") Then
        '        PLC_UPS_Battery_Backup = 1
        '    Else
        '        PLC_UPS_Battery_Backup = 0
        '    End If
        'Else
        '    PLC_UPS_Battery_Backup = 0
        'End If

        'If Not ADA_Setup.amp_mon_box.SelectedItem Is Nothing Then
        '    'Amp Mon
        '    If String.Equals(ADA_Setup.amp_mon_box.SelectedItem.ToString, "Yes") Then
        '        Amp_Mon_Electro = 1
        '    Else
        '        Amp_Mon_Electro = 0
        '    End If
        'Else
        '    Amp_Mon_Electro = 0
        'End If

        'If Not ADA_Setup.remote_box.SelectedItem Is Nothing Then
        '    'remote io
        '    If String.Equals(ADA_Setup.remote_box.SelectedItem.ToString, "Yes") Then
        '        Remote_IO = 1
        '    Else
        '        Remote_IO = 0
        '    End If
        'Else
        '    Remote_IO = 0
        'End If

        '---------------------------------------------------------------------------------------
        Me.load_device_Type_specs()  'load dropboxes
        Me.Load_textboxes_inputs()  'load textbox inputs from device type and disconnect


        '---------------- Fill the dictionary ADA parts list ------------------------
        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT distinct ADA_Number from ada_active_parts where ada_number is not null"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            'fill list with ada numbers and qty 0.
            If reader.HasRows Then
                While reader.Read
                    If Starter_Panel.ContainsKey(reader(0)) = False Then
                        Starter_Panel.Add(reader(0), 0)
                        IO_Panel.Add(reader(0), 0)
                        PLC_Panel.Add(reader(0), 0)
                        Field.Add(reader(0), 0)
                        Scanners.Add(reader(0), 0)
                        '  M12.Add(reader(0), 0)
                        '   M12_ES.Add(reader(0), 0)
                    End If
                End While
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Property myName() As String

        Get
            Return ADA_set_name
        End Get

        Set(value As String)
            ADA_set_name = value
        End Set

    End Property

    '-------------------  This method will update the values of the inputs array  --------------------------

    Public Sub Update_inputs(motor_data As Array, inputs_data As Array, push_data As Array, lights_data As Array, field_data As Array, bus_data As Array, scanner_data As Array, brakes_data As Array)
        Me.Motors_table = motor_data.Clone
        Me.inputs_table = inputs_data.Clone
        Me.push_table = push_data.Clone
        Me.lights_table = lights_data.Clone
        Me.field_table = field_data.Clone
        Me.bus_table = bus_data.Clone
        Me.scanner_table = scanner_data.Clone
        Me.brakes_table = brakes_data.Clone
    End Sub


    '------------------ This method will update global settings variable of the set --------------------------

    Public Sub Update_globals(Voltage, efficiency, Power_factor, include_ADA_brackets, Use_4pts_IO, Use_8pts_IO, Use_16pts_IO, removal_msbox_t, M23_at_IO_Receptacle,
            M23_Bulkhead_at_ADA, Use_4pts_IO_RB, Use_8pts_IO_RB, Use_16pts_IO_RB_Inputs, Use_16pts_IO_RB_Outputs, Single_Channel_IO_1_Input_per_RB, Single_Channel_IO_1_Output_per_RB, Splitter_Perc_dual_ch, Valve_solenoid_adapter, Smart_Wire_Darwin_Limited,
            Smart_Wire_Darwin_Full, EthernetIP_Full, Percent_overage_inputs, Percent_overage_nm_outputs, Percent_overage_m_outputs)

        Me.Voltage = Voltage
        Me.efficiency = efficiency
        Me.Power_factor = Power_factor
        Me.include_ADA_brackets = include_ADA_brackets
        Me.Use_4pts_IO = Use_4pts_IO
        Me.Use_8pts_IO = Use_8pts_IO
        Me.Use_16pts_IO = Use_16pts_IO
        Me.removal_msbox = removal_msbox_t
        Me.M23_at_IO_Receptacle = M23_at_IO_Receptacle
        Me.M23_Bulkhead_at_ADA = M23_Bulkhead_at_ADA
        Me.Use_4pts_IO_RB = Use_4pts_IO_RB
        Me.Use_8pts_IO_RB = Use_8pts_IO_RB
        Me.Use_16pts_IO_RB_Inputs = Use_16pts_IO_RB_Inputs
        Me.Use_16pts_IO_RB_Outputs = Use_16pts_IO_RB_Outputs
        Me.Single_Channel_IO_1_Input_per_RB = Single_Channel_IO_1_Input_per_RB
        Me.Single_Channel_IO_1_Output_per_RB = Single_Channel_IO_1_Output_per_RB
        Me.Splitter_Perc_dual_ch = Splitter_Perc_dual_ch
        Me.Valve_solenoid_adapter = Valve_solenoid_adapter
        Me.Smart_Wire_Darwin_Limited = Smart_Wire_Darwin_Limited
        Me.Smart_Wire_Darwin_Full = Smart_Wire_Darwin_Full
        Me.EthernetIP_Full = EthernetIP_Full
        Me.Percent_overage_inputs = Percent_overage_inputs
        Me.Percent_overage_nm_outputs = Percent_overage_nm_outputs
        Me.Percent_overage_m_outputs = Percent_overage_m_outputs

        If String.Equals(Voltage, "230") = True Then
            voltage_index = 0
        ElseIf String.Equals(Voltage, "480") = True Then
            voltage_index = 1
        Else
            voltage_index = 2
        End If


    End Sub


    '-- Load all inputs from textboxes (device type and disconnects) to set variables
    Public Sub Load_textboxes_inputs()

        '  Me.R2_setup = If((String.IsNullOrEmpty(ADA_Setup.r2_setup_text.Text) = False And IsNumeric(ADA_Setup.r2_setup_text.Text) = True), (ADA_Setup.r2_setup_text.Text), 0)   'r2_setup
        '     Me.Amp_30_Drops = If((String.IsNullOrEmpty(ADA_Setup.amp_30_text.Text) = False And IsNumeric(ADA_Setup.amp_30_text.Text) = True), (ADA_Setup.amp_30_text.Text), 0)   'Amp_30_Drops 
        Me.ADA_NEMA12_BOX24_No_Starters = If((String.IsNullOrEmpty(ADA_Setup.ada_nema12box24_text.Text) = False And IsNumeric(ADA_Setup.ada_nema12box24_text.Text) = True), (ADA_Setup.ada_nema12box24_text.Text), 0)   'ADA_NEMA12_BOX24_No_Starters
        Me.ADA_NEMA12_BOX30_No_Starters = If((String.IsNullOrEmpty(ADA_Setup.ada_nema12box30_text.Text) = False And IsNumeric(ADA_Setup.ada_nema12box30_text.Text) = True), (ADA_Setup.ada_nema12box30_text.Text), 0)   '  ADA_NEMA12_BOX30_No_Starters
        Me.motor_disconnect = If((String.IsNullOrEmpty(ADA_Setup.disconnect_text.Text) = False And IsNumeric(ADA_Setup.disconnect_text.Text) = True), (ADA_Setup.disconnect_text.Text), 0)   'motor_disconnect
        Me.MDRs = If((String.IsNullOrEmpty(ADA_Setup.MDRs.Text) = False And IsNumeric(ADA_Setup.MDRs.Text) = True), (ADA_Setup.MDRs.Text), 0) 'MDRs motor driven rollers

    End Sub

    '-- load comboboxes from device tab panel
    Public Sub load_device_Type_specs()

        '--------------- Boolean values in Device Type specifications -----------------------

        Panel_bool = 1  'default panel present



        If Not ADA_Setup.Nema12_com.SelectedItem Is Nothing Then
            'NEMA 12 and 4x
            If String.Equals(ADA_Setup.Nema12_com.SelectedItem.ToString, "NEMA12") Then
                NEMA12 = 1
                NEMA4X = 0
            Else
                NEMA12 = 0
                NEMA4X = 1
            End If
        Else
            NEMA12 = 1
            NEMA4X = 0
        End If

        'If Not ADA_Setup.nema4x_box.SelectedItem Is Nothing Then
        '    'NEMA 4x
        '    If String.Equals(ADA_Setup.nema4x_box.SelectedItem.ToString, "Yes") Then
        '        NEMA4X = 1
        '    Else
        '        NEMA4X = 0
        '    End If
        'Else
        '    NEMA4X = 0
        'End If

        If Not ADA_Setup.PLC_gate_box.SelectedItem Is Nothing Then
            'PLC Gateway
            If String.Equals(ADA_Setup.PLC_gate_box.SelectedItem.ToString, "Yes") Then
                PLC_GateWay_PGW = 1
            Else
                PLC_GateWay_PGW = 0
            End If
        Else
            PLC_GateWay_PGW = 0
        End If

        If Not ADA_Setup.control_box.SelectedItem Is Nothing Then
            'Control logic
            If String.Equals(ADA_Setup.control_box.SelectedItem.ToString, "Yes") Then
                ControlLogix_PLC_Box = 1
            Else
                ControlLogix_PLC_Box = 0
            End If
        Else
            ControlLogix_PLC_Box = 0
        End If

        If Not ADA_Setup.compact_box.SelectedItem Is Nothing Then
            'Compact logic
            If String.Equals(ADA_Setup.compact_box.SelectedItem.ToString, "Yes") Then
                CompactLogix_PLC_Box = 1
            Else
                CompactLogix_PLC_Box = 0
            End If
        Else
            CompactLogix_PLC_Box = 0
        End If

        If Not ADA_Setup.PLC_UPS_box.SelectedItem Is Nothing Then
            'PLC UPS
            If String.Equals(ADA_Setup.PLC_UPS_box.SelectedItem.ToString, "Yes") Then
                PLC_UPS_Battery_Backup = 1
            Else
                PLC_UPS_Battery_Backup = 0
            End If
        Else
            PLC_UPS_Battery_Backup = 0
        End If

        If Not ADA_Setup.amp_mon_box.SelectedItem Is Nothing Then
            'Amp Mon
            If String.Equals(ADA_Setup.amp_mon_box.SelectedItem.ToString, "Yes") Then
                Amp_Mon_Electro = 1
            Else
                Amp_Mon_Electro = 0
            End If
        Else
            Amp_Mon_Electro = 0
        End If

        If Not ADA_Setup.remote_box.SelectedItem Is Nothing Then
            'remote io
            If String.Equals(ADA_Setup.remote_box.SelectedItem.ToString, "Yes") Then
                Remote_IO = 1
            Else
                Remote_IO = 0
            End If
        Else
            Remote_IO = 0
        End If

    End Sub



    '---------------------------------------  This will calculate the ADA parts in the Starter Panel (Starter Panel tab) ---------------------------------

    Public Sub Calculate_Starter_Panel(Motor_horsepower(,) As Double, VFD_horsepower(,) As Double, Starters_values() As Double, value_to_index_480() As Integer, value_to_index_575() As Integer, Power_supply_480(,) As Double, Power_supply_575(,) As Double)

        '----- This section calculate the secondary values at the bottom of the starter panel tab necessary to get qty of ada parts-------


        '------------ Automated inputs in setup field ( F93, F170, F172, F173, F179, F180 )

        Dim n_field As Double : n_field = Math.Round(If(Panel_bool > 0, ((CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8) + ((CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 4), 0) + (CType(Me.field_table(9, 0), Double) + CType(Me.field_table(10, 0), Double)) / 24, MidpointRounding.AwayFromZero)

        Me.inputs_table(18, 0) = Me.inputs_table(12, 0)  'make sure ADA photoeye brackets qty is the same as Photoeye 18mm Mnt  (F93)

        Me.brakes_table(4, 0) = Math.Round(0.08 * (CType(Me.field_table(1, 0), Double) + CType(Me.field_table(2, 0), Double) + CType(Me.field_table(3, 0), Double)), MidpointRounding.AwayFromZero) 'F170
        Me.brakes_table(6, 0) = Math.Round(CType(Me.brakes_table(4, 0), Double) + CType(Me.brakes_table(5, 0), Double), MidpointRounding.AwayFromZero)  'F172
        Me.brakes_table(7, 0) = Math.Round(If(n_field > 0, n_field / 30, 0) + (CType(Me.brakes_table(6, 0), Double) / 120), MidpointRounding.AwayFromZero) 'F173
        Me.brakes_table(12, 0) = Math.Round(CType(Me.brakes_table(8, 0), Double) + (CType(Me.brakes_table(9, 0), Double) * 20) + (8 * CType(Me.brakes_table(10, 0), Double) + CType(Me.brakes_table(11, 0), Double)), MidpointRounding.AwayFromZero)  'F179
        Me.brakes_table(13, 0) = If((MDRs + ((CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)) * 2)) >= 60, Math.Ceiling((MDRs / 60) + ((CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 30)), 0) 'F180




        'Initialize result_motor_amps and result_vfd_amps
        For i = 0 To 11
            For j = 0 To 15
                Results_amps_motors(i, j) = 0 : Results_amps_vfd(i, j) = 0
            Next
        Next

        '------------------------------------------------------------------
        'Starters for non reversing/reversing motors
        For i = 0 To 11
            For j = 0 To 15
                If Motor_horsepower(i, voltage_index) * 0.85 * 1.25 <= Starters_values(j) Then
                    Results_amps_motors(i, j) = 1
                    Exit For
                End If
            Next
        Next


        'Starters for VFDs
        For i = 0 To 11
            For j = 0 To 15
                If VFD_horsepower(i, voltage_index) <= Starters_values(j) Then
                    Results_amps_vfd(i, j) = 1
                    Exit For
                End If
            Next
        Next

        '//////////////////////////--------------------------------------------------------------------- (non reversing)
        Dim non_reversing_qty(11) As Double
        Dim non_reversing_amps As Double : non_reversing_amps = 0
        Dim non_reversing_starter_qty(14) As Double

        For i = 0 To 11
            non_reversing_qty(i) = Panel_bool * CType(Me.Motors_table(i, 0), Double)
        Next
        non_reversing_qty(4) = Panel_bool * (CType(Me.field_table(7, 0), Double) + CType(Me.Motors_table(4, 0), Double))  '1.50 HP non-reversing HP

        For i = 0 To 11
            non_reversing_amps = non_reversing_amps + (non_reversing_qty(i) * Motor_horsepower(i, voltage_index))
        Next

        For i = 0 To 14
            If (i < 5) Then
                For j = 0 To 9
                    non_reversing_starter_qty(i) = non_reversing_starter_qty(i) + (non_reversing_qty(j) * Results_amps_motors(j, i + 1))
                Next
            Else
                For j = 0 To 11
                    non_reversing_starter_qty(i) = non_reversing_starter_qty(i) + (non_reversing_qty(j) * Results_amps_motors(j, i + 1))
                Next
            End If

        Next

        '//////////////////////////////------------------------------------------------------------------ (reversing)
        Dim reversing_qty(11) As Integer
        Dim reversing_amps As Double
        Dim reversing_starter_qty(14) As Integer

        For i = 0 To 11
            reversing_qty(i) = Panel_bool * CType(Me.Motors_table(i + 12, 0), Double)
        Next

        For i = 0 To 11
            reversing_amps = reversing_amps + (reversing_qty(i) * Motor_horsepower(i, voltage_index))
        Next

        For i = 0 To 14
            If (i < 7) And i <> 0 Then
                For j = 0 To 9
                    reversing_starter_qty(i) = reversing_starter_qty(i) + (reversing_qty(j) * Results_amps_motors(j, i + 1))
                Next
            Else
                For j = 0 To 11
                    reversing_starter_qty(i) = reversing_starter_qty(i) + (reversing_qty(j) * Results_amps_motors(j, i + 1))
                Next
            End If
        Next

        '/////////////////////////////////------------------------------------ (VFD)

        Dim VFD_qty(11) As Integer          'VFD quantities 
        Dim VFD_amps As Double
        Dim VFD_starter_qty(14) As Integer  'VFD Starter Quantities

        For i = 0 To 11
            VFD_qty(i) = Panel_bool * CType(Me.Motors_table(i + 24, 0), Double)
        Next

        For i = 0 To 11
            VFD_amps = VFD_amps + (VFD_qty(i) * VFD_horsepower(i, voltage_index))
        Next


        For i = 0 To 14
            If (i < 9) And i <> 0 Then
                For j = 0 To 9
                    VFD_starter_qty(i) = VFD_starter_qty(i) + (VFD_qty(j) * Results_amps_vfd(j, i + 1))
                Next
            Else
                For j = 0 To 11
                    VFD_starter_qty(i) = VFD_starter_qty(i) + (VFD_qty(j) * Results_amps_vfd(j, i + 1))
                Next
            End If
        Next


        '//////////////////////////---------------------------------
        Dim VFD_Q(6) As Integer    'VFD Quantities (with uppercase lol)

        VFD_Q(0) = VFD_qty(0) + VFD_qty(1) : VFD_Q(1) = VFD_qty(2) + VFD_qty(3) : VFD_Q(2) = VFD_qty(4) + VFD_qty(5) : VFD_Q(3) = VFD_qty(6) + VFD_qty(7)
        VFD_Q(4) = VFD_qty(8) + VFD_qty(9) : VFD_Q(5) = VFD_qty(10) : VFD_Q(6) = VFD_qty(11)


        '///////////////////////////--------------------------------------- (Limited function soft)

        Dim Limited_VFD_Soft_qty(11) As Integer  'Limited Function VFDs Soft Start Quantities
        Dim Soft_start_amps As Double
        Dim Limited_VFD_Soft_MS(14) As Integer  'Limited Function VFDs Soft Start M/S Starter Quantities

        For i = 0 To 11
            Limited_VFD_Soft_qty(i) = Panel_bool * CType(Me.Motors_table(i + 36, 0), Double)
        Next

        For i = 0 To 11
            Soft_start_amps = Soft_start_amps + (Limited_VFD_Soft_qty(i) * Motor_horsepower(i, voltage_index))
        Next


        For i = 0 To 14
            If (i < 9) Then
                For j = 0 To 9
                    Limited_VFD_Soft_MS(i) = Limited_VFD_Soft_MS(i) + (Limited_VFD_Soft_qty(j) * Results_amps_vfd(j, i + 1))
                Next
            Else
                For j = 0 To 11
                    Limited_VFD_Soft_MS(i) = Limited_VFD_Soft_MS(i) + (Limited_VFD_Soft_qty(j) * Results_amps_vfd(j, i + 1))
                Next
            End If
        Next

        '////////////////////////--------------------
        Dim Limited_VFD_Soft_Q(6) As Integer  'Limited Function VFDs Soft Start Quantities

        Limited_VFD_Soft_Q(0) = Limited_VFD_Soft_qty(0) + Limited_VFD_Soft_qty(1) : Limited_VFD_Soft_Q(1) = Limited_VFD_Soft_qty(2) + Limited_VFD_Soft_qty(3) : Limited_VFD_Soft_Q(2) = Limited_VFD_Soft_qty(4) + Limited_VFD_Soft_qty(5)
        Limited_VFD_Soft_Q(3) = Limited_VFD_Soft_qty(7) : Limited_VFD_Soft_Q(4) = Limited_VFD_Soft_qty(8) + Limited_VFD_Soft_qty(9) : Limited_VFD_Soft_Q(5) = Limited_VFD_Soft_qty(9) + Limited_VFD_Soft_qty(10) : Limited_VFD_Soft_Q(6) = Limited_VFD_Soft_qty(10) + Limited_VFD_Soft_qty(11)

        '-------- amps -------------
        non_reversing_amps_p = non_reversing_amps
        reversing_amps_p = reversing_amps
        vfd_amps_p = VFD_amps
        soft_amps_p = Soft_start_amps


        ' ----------------------------------------------------------------------------------------------

        Dim panel_width As Double : panel_width = 350  'panel width  (406.4 actual)
        Panel_Space_Top = 0
        Panel_Space_Bottom = 0


        'DC amps on main PS
        DC_amps_main_PS = Math.Round(Panel_bool * (4.8 + (5 * ControlLogix_PLC_Box) + (4 * CompactLogix_PLC_Box) + CType(Me.field_table(8, 0), Double) + CType(Me.brakes_table(6, 0), Double)))

        'DC amps on auxiliary PS
        DC_amps_aux_PS = Panel_bool * CType(Me.brakes_table(12, 0), Double)


        '---------- Aux --------

        '///////////////// Aux PS  /////////////////
        Aux_PS_5A = 0
        Aux_PS_10A = 0
        Aux_PS_20A = 0
        Aux_PS_40A = 0

        If voltage_index < 2 Then
            If DC_amps_aux_PS <= 420 Then
                For j = 0 To 50
                    If value_to_index_480(j) <= DC_amps_aux_PS Then
                        Aux_PS_5A = Power_supply_480(j, 0)
                        Aux_PS_10A = Power_supply_480(j, 1)
                        Aux_PS_20A = Power_supply_480(j, 2)
                        Aux_PS_40A = Power_supply_480(j, 3)
                        Exit For
                    End If
                Next
            End If

        Else

            If DC_amps_aux_PS <= 270 Then
                For j = 0 To 50
                    If value_to_index_575(j) <= DC_amps_aux_PS Then
                        Aux_PS_5A = Power_supply_575(j, 0)
                        Aux_PS_10A = Power_supply_575(j, 1)
                        Aux_PS_20A = Power_supply_575(j, 2)
                        Aux_PS_40A = Power_supply_575(j, 3)
                        Exit For
                    End If
                Next
            End If
        End If
        '//////////////////////////////////////////////


        '-------------------------------------------------------------------------------
        ' Non Reversing Starter (Starter Panel tab) cell (H319)

        non_reversing_starters = 0
        For i = 0 To 11
            non_reversing_starters = non_reversing_starters + non_reversing_qty(i)
        Next
        non_reversing_starters = (non_reversing_starters) * Panel_bool + (Aux_PS_5A + Aux_PS_10A + Aux_PS_20A + Aux_PS_40A)


        '---------------------------------------------------------------------------------------
        ' Reversing Starter (Starter Panel tab) cell (H320)
        reversing_starters = 0
        For i = 0 To 11
            reversing_starters = reversing_starters + reversing_qty(i)
        Next
        reversing_starters = reversing_starters * Panel_bool
        '------------------------------------------------------------------------------------------------

        ' # VFD / Starter (Starter Panel tab) cell (H321)
        VFD_starters = 0
        For i = 0 To 11
            VFD_starters = VFD_starters + VFD_qty(i)
        Next
        VFD_starters = VFD_starters * Panel_bool
        '------------------------------------------------------------------------------------------------

        ' # Soft VFD Starts (Starter Panel tab) cell (H322)
        Soft_starters = 0
        For i = 0 To 11
            Soft_starters = Soft_starters + Limited_VFD_Soft_qty(i)
        Next
        Soft_starters = Soft_starters * Panel_bool
        '---------------------------------------------------------------------------------------

        ' # Starters
        Starters = non_reversing_starters + reversing_starters + VFD_starters + Soft_starters

        '-------------------------------------------------------------------------------------

        '# starter panels and #big panels

        big_panels = 0

        '------------------------------------------------------------------
        '# PLC
        PLC = Panel_bool * (ControlLogix_PLC_Box + CompactLogix_PLC_Box)

        '--------------------------------------------------------------------------------
        'Total current per group

        Total_current_per_group = 0 'initialize this variable, it is declare later

        '-------------------------------------------------------------------------------

        'Primary PS 5A
        Primary_PS_5A = If((DC_amps_main_PS > 0 And DC_amps_main_PS <= 5), 1, 0)
        Primary_PS_10A = If((DC_amps_main_PS > 5 And DC_amps_main_PS <= 10), 1, 0)
        Primary_PS_20A = If((DC_amps_main_PS > 10 And DC_amps_main_PS <= 20), 1, 0)
        Primary_PS_40A = If((DC_amps_main_PS > 20 And DC_amps_main_PS <= 40), 1, 0) + If((DC_amps_main_PS > 40 And DC_amps_main_PS <= 80), 2, 0)

        'If Primary_PS_40A < 0 Then MessageBox.Show("Too many amps requested on primary PS")



        For i = 1 To 6  'do this selection case twice to make sure all qtys were calculated

            For Each kvp As KeyValuePair(Of String, Double) In Starter_Panel.ToArray

                Select Case kvp.Key
                    Case "ADA-2PS-05-24"
                        Starter_Panel(kvp.Key) = Math.Round(Aux_PS_5A, MidpointRounding.AwayFromZero)
                    Case "ADA-2PS-10-24"
                        Starter_Panel(kvp.Key) = Math.Round(Primary_PS_5A + Primary_PS_10A + Aux_PS_10A, MidpointRounding.AwayFromZero)
                    Case "ADA-2PS-20-24"
                        Starter_Panel(kvp.Key) = Math.Round(Primary_PS_20A + Aux_PS_20A, MidpointRounding.AwayFromZero)
                    Case "ADA-2PS-40-24"
                        Starter_Panel(kvp.Key) = Math.Round((Primary_PS_40A + Aux_PS_40A) + (3 * (CType(Me.brakes_table(7, 0), Double))) + (3 * (CType(Me.brakes_table(13, 0), Double))), MidpointRounding.AwayFromZero)
                    Case "ADA-120VAC-2.0KVA-TRANS"
                        Starter_Panel(kvp.Key) = Math.Ceiling(CType(Me.brakes_table(0, 0), Double) / 40) * Panel_bool + CType(Me.brakes_table(3, 0), Double)
                    Case "ADA-3CB-30"
                        Starter_Panel(kvp.Key) = Math.Round((Panel_bool * big_panels) + ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters + CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double), MidpointRounding.AwayFromZero)
                    Case "ADA-3CB-RED_HANDLE"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-3CB-30")
                    Case "ADA-3CB-30-MULTITAP"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro > 0, (Panel_bool * big_panels) + ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters, 0), MidpointRounding.AwayFromZero)



                            '------------------------------ Motor Starters --------------------------------
                    Case "ADA-3PMC-2.5-4.0-6.3-10"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_qty(8) + VFD_qty(8), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3PMC-10PLUS"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_qty(9) + VFD_qty(9), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3PMS-2.5-4.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, VFD_starter_qty(8) + CType(Me.brakes_table(13, 0), Double), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3PMS-4-6.3"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, VFD_starter_qty(9) + CType(Me.brakes_table(13, 0), Double), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3PMS-6.3-10"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, VFD_starter_qty(10) + CType(Me.brakes_table(13, 0), Double), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3PMS-10-16"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, VFD_starter_qty(11) + CType(Me.brakes_table(13, 0), Double), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-0.1-0.16"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(0) + VFD_starter_qty(0), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-0.16-0.25"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(1) + VFD_starter_qty(1), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-0.25-0.4"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(2) + VFD_starter_qty(2) + Limited_VFD_Soft_MS(2), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS- 0.4-0.63"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(3) + VFD_starter_qty(3) + Limited_VFD_Soft_MS(3), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-0.63-1.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(4) + VFD_starter_qty(4) + Limited_VFD_Soft_MS(4), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-1.0-1.6"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, (non_reversing_starter_qty(5) + VFD_starter_qty(5) + Limited_VFD_Soft_MS(5) + CType(Me.brakes_table(2, 0), Double) + Aux_PS_5A + Aux_PS_10A + Aux_PS_20A), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-1.6-2.5"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(6) + VFD_starter_qty(6) + Limited_VFD_Soft_MS(6) + Aux_PS_40A, 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-2.5-4.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(7) + Limited_VFD_Soft_MS(7), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-4.0-6.3"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(8) + Limited_VFD_Soft_MS(8) + CType(Me.brakes_table(13, 0), Double), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-6.3-10.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(9) + Limited_VFD_Soft_MS(9), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-8-12"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(10) + Limited_VFD_Soft_MS(10), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-10-16"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(11) + Limited_VFD_Soft_MS(11), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-16-20"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(12) + Limited_VFD_Soft_MS(12), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-20-25"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(13) + VFD_starter_qty(13) + Limited_VFD_Soft_MS(13), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3MS-25-32"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, non_reversing_starter_qty(14) + VFD_starter_qty(14) + Limited_VFD_Soft_MS(14), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-0.1-0.16"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(0), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-0.16-0.25"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(1), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-0.25-0.4"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(2), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS- 0.4-0.63"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(3), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-0.63-1.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(4), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-1.0-1.6"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(5), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-1.6-2.5"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(6), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-2.5-4.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(7), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-4.0-6.3"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(8), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-6.3-10.0"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(9), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-8-12"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(10), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-10-16"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(11), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-16-20"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(12), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-3RMS-20-25"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, reversing_starter_qty(13), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-0.5-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(0), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-1-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(1), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-2-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(2), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-3-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(3), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-5-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(4), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-7.5-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(5), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-10-AB-575"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 2, VFD_Q(6), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-0.5-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(0), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-1-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(1), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-2-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(2), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-3-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(3), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-5-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(4), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-7.5-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(5), 0), MidpointRounding.AwayFromZero)
                    Case "ADA-VFD-10-AB-525"
                        Starter_Panel(kvp.Key) = Math.Round(If(voltage_index = 1, VFD_Q(6), 0), MidpointRounding.AwayFromZero)


                        '---------- B-Coded M12 Cables ---------------------------------------------
                    Case "ADA-SAC5-02-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-03-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-04-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-05-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-06-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-10-B"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SAC5-15-B"
                        Starter_Panel(kvp.Key) = 0



                        '--------------- Terminals ---------------------------------------------

                    Case "ADA-1TEA"
                        Starter_Panel(kvp.Key) = Math.Ceiling(((CType(Me.brakes_table(7, 0), Double)) * 6) + ((CType(Me.brakes_table(7, 0), Double)) * 4) + (6 * (CType(Me.brakes_table(13, 0), Double))) + (4 * (CType(Me.Motors_table(24, 0), Double) +
                            CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) _
                            + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double))))
                    Case "ADA-1TB2_2"
                        Starter_Panel(kvp.Key) = Math.Ceiling(((CType(Me.brakes_table(7, 0), Double)) * 30) + (4 * (CType(Me.Motors_table(24, 0), Double) +
                            CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) _
                            + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double))))
                    Case "ADA-1EB2"
                        Starter_Panel(kvp.Key) = Math.Ceiling(((CType(Me.brakes_table(7, 0), Double)) * 6) + (4 * (CType(Me.Motors_table(24, 0), Double) +
                            CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) _
                            + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double))))
                    Case "ADA-1TB2-JP2"
                        Starter_Panel(kvp.Key) = Math.Ceiling(((CType(Me.brakes_table(7, 0), Double)) * 4) + (4 * (CType(Me.Motors_table(24, 0), Double) +
                            CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) _
                            + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double))))
                    Case "ADA-1TB2-JP4"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-1TB2-JP2")
                    Case "ADA-1TB2-JP10"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-1TB2-JP2")
                    Case "ADA-1TB4"
                        Starter_Panel(kvp.Key) = 0 '2 * Panel_bool * (Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24"))
                    Case "ADA-1EB4"
                        Starter_Panel(kvp.Key) = 2 * Panel_bool
                    Case "ADA-1TB6"
                        Starter_Panel(kvp.Key) = (CType(Me.brakes_table(9, 0), Double) * 4) + (CType(Me.brakes_table(7, 0), Double) * 21) + (CType(Me.brakes_table(3, 0), Double) * 2) + (CType(Me.brakes_table(13, 0), Double) * 14) + (6 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters)) + (Panel_bool * (Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24")))
                    Case "ADA-1EB6"
                        Starter_Panel(kvp.Key) = 4 * Panel_bool + (CType(Me.brakes_table(3, 0), Double))
                    Case "ADA-1TB6-JP2"
                        Starter_Panel(kvp.Key) = Panel_bool * (2 * CType(Me.brakes_table(9, 0), Double)) + (9 * CType(Me.brakes_table(7, 0), Double)) + (3 * CType(Me.brakes_table(13, 0), Double)) + (CType(Me.brakes_table(7, 0), Double)) * 3 + (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters) * 3
                    Case "ADA-1TB6-JP5"
                        Starter_Panel(kvp.Key) = Panel_bool * 4 + CType(Me.brakes_table(4, 0), Double)
                    Case "ADA-1TB10"
                        Starter_Panel(kvp.Key) = ((CType(Me.brakes_table(9, 0), Double)) * 2 * Panel_bool) + (9 * (CType(Me.brakes_table(7, 0), Double))) + (14 * (CType(Me.brakes_table(13, 0), Double))) + (6 * (CType(Me.brakes_table(7, 0), Double))) + (6 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-1TB10-JP2"
                        Starter_Panel(kvp.Key) = ((CType(Me.brakes_table(9, 0), Double)) * 2 * Panel_bool) + (9 * (CType(Me.brakes_table(7, 0), Double))) + (3 * (CType(Me.brakes_table(13, 0), Double))) + (3 * (CType(Me.brakes_table(7, 0), Double))) + (3 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-1TB10-JP5"
                        Starter_Panel(kvp.Key) = ((CType(Me.brakes_table(9, 0), Double)) * 2 * Panel_bool) + (9 * (CType(Me.brakes_table(7, 0), Double))) + (2 * (CType(Me.brakes_table(13, 0), Double)))
                    Case "ADA-1EB10"
                        Starter_Panel(kvp.Key) = ((CType(Me.brakes_table(9, 0), Double)) * Panel_bool) + (9 * (CType(Me.brakes_table(7, 0), Double))) + (5 * (CType(Me.brakes_table(13, 0), Double))) + (4 * (CType(Me.brakes_table(7, 0), Double))) + (1 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-6x10-W-TLBL"
                        Starter_Panel(kvp.Key) = Panel_bool * (Starter_Panel.Item("ADA-1TB4") + Starter_Panel.Item("ADA-1EB4") + Starter_Panel.Item("ADA-1TB6") + Starter_Panel.Item("ADA-1EB6") + Starter_Panel.Item("ADA-1TB10"))
                    Case "ADA-RJ45-SWT-8-port"
                        Starter_Panel(kvp.Key) = Math.Round((CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) _
                            + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 4)
                    Case "ADA-RJ45-SWT-16-port"
                        Starter_Panel(kvp.Key) = 0

                        ' ------------------------------- Safety Relay --------------------------

                    Case "ADA-1SRB"
                        Starter_Panel(kvp.Key) = (CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))

                    Case "ADA-1SRY"
                        Starter_Panel(kvp.Key) = (CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))


                        '----------------------------- Motor Starter Communication ---------------------

                    Case "ADA-SWD-CM"
                        Starter_Panel(kvp.Key) = Math.Round(Panel_bool * ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2))) * removal_msbox
                    Case "ADA-SWD-PFM-8"
                        Starter_Panel(kvp.Key) = Math.Round((Panel_bool * Starters + Starter_Panel.Item("ADA-SWD-4DI-2DO")) / 14, MidpointRounding.AwayFromZero) * removal_msbox
                    Case "ADA-EIP-CM-MS"
                        Starter_Panel(kvp.Key) = (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))
                    Case "ADA-IOL-CM-MS"
                        Starter_Panel(kvp.Key) = ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8
                    Case "ADA-SWD-4DI-2DO"
                        Starter_Panel(kvp.Key) = If((Smart_Wire_Darwin_Full + EthernetIP_Full) = 0, Math.Ceiling((VFD_Q(0) + VFD_Q(1) + VFD_Q(2) + VFD_Q(3) + VFD_Q(4) + VFD_Q(5) + VFD_Q(6)) / 2), 0) + If(Smart_Wire_Darwin_Full = 1, CType(Me.brakes_table(0, 0), Double) + CType(Me.brakes_table(1, 0), Double), 0) * removal_msbox
                    Case "ADA-SWD-RIB-DEVICE-PLUG"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-SWD-CM") + Starter_Panel.Item("ADA-SWD-PFM-8") + Starter_Panel.Item("ADA-SWD-4DI-2DO")) + (Limited_VFD_Soft_Q(0) + Limited_VFD_Soft_Q(1) + Limited_VFD_Soft_Q(2) + Limited_VFD_Soft_Q(3) + Limited_VFD_Soft_Q(4) + Limited_VFD_Soft_Q(5) + Limited_VFD_Soft_Q(6))
                    Case "ADA-SWD-FF-Coupling"
                        Starter_Panel(kvp.Key) = If((big_panels - (CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double)) * Panel_bool) > 1, (big_panels * Panel_bool) - 1, 0)
                    Case "ADA-SWD000-RIB-END-PLUG"
                        Starter_Panel(kvp.Key) = (big_panels * Panel_bool) - (CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double))
                    Case "ADA-SWD4-8FRF-10"
                        Starter_Panel(kvp.Key) = Math.Round(((CType(Me.Motors_table(0, 0), Double)) + (CType(Me.Motors_table(1, 0), Double)) + (CType(Me.Motors_table(2, 0), Double)) + (CType(Me.Motors_table(3, 0), Double)) + (CType(Me.Motors_table(4, 0), Double)) + (CType(Me.Motors_table(5, 0), Double)) + (CType(Me.Motors_table(6, 0), Double)) _
                            + (CType(Me.Motors_table(7, 0), Double)) + (CType(Me.Motors_table(8, 0), Double)) + (CType(Me.Motors_table(9, 0), Double)) + (CType(Me.Motors_table(10, 0), Double)) + (CType(Me.Motors_table(11, 0), Double)) + (CType(Me.Motors_table(12, 0), Double)) _
                            + (CType(Me.Motors_table(13, 0), Double)) + (CType(Me.Motors_table(14, 0), Double)) + (CType(Me.Motors_table(15, 0), Double)) + (CType(Me.Motors_table(16, 0), Double)) + (CType(Me.Motors_table(17, 0), Double)) _
                            + (CType(Me.Motors_table(18, 0), Double)) + (CType(Me.Motors_table(19, 0), Double)) + (CType(Me.Motors_table(20, 0), Double)) + (CType(Me.Motors_table(21, 0), Double)) + (CType(Me.Motors_table(22, 0), Double)) + (CType(Me.Motors_table(23, 0), Double))) / 6) * 2 * removal_msbox
                    Case "ADA-SWD8-TERM"
                        Starter_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-SWD005-RIB"
                        Starter_Panel(kvp.Key) = (big_panels - (CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double))) * (Panel_bool / 2)




                        '------------------------------ Wiring -------------------------
                    Case "ADA-SAC5-0,5-G-PWR-Bulkhead"
                        Starter_Panel(kvp.Key) = ((CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))) * 2 'Math.Round(CType(Me.bus_table(9, 0), Double) / 8, MidpointRounding.AwayFromZero)

                    Case "ADA-SAC5-0,5-G-PWR-M-Bulkhead"
                        Starter_Panel(kvp.Key) = ((CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))) * 2

                    Case "ADA-EIP-Field-Patch"
                        Starter_Panel(kvp.Key) = (CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))
                    Case "ADA-EIP_M12-RJ45-Bulkhead"
                        Starter_Panel(kvp.Key) = ((CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))) * 2
                    Case "ADA-BHR5-M-PG-9"
                        Starter_Panel(kvp.Key) = (CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))
                    Case "ADA-BHR5-F-PG-9"
                        Starter_Panel(kvp.Key) = (CType(Me.Motors_table(24, 0), Double) + CType(Me.Motors_table(25, 0), Double) + CType(Me.Motors_table(26, 0), Double) + CType(Me.Motors_table(27, 0), Double) + CType(Me.Motors_table(28, 0), Double) + CType(Me.Motors_table(29, 0), Double) + CType(Me.Motors_table(30, 0), Double) + CType(Me.Motors_table(31, 0), Double) _
                            + CType(Me.Motors_table(32, 0), Double) + CType(Me.Motors_table(33, 0), Double) + CType(Me.Motors_table(34, 0), Double) + CType(Me.Motors_table(35, 0), Double)) / 2 + (If(removal_msbox = 0, ((Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32")) + ((Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) * 2)) * Panel_bool / 8, 0))
                    Case "ADA-Hytrol MDR PWR CBL"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-Hytrol MDR EXT CBL"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-RAW MDR  CBL"
                        Starter_Panel(kvp.Key) = Math.Ceiling(2 * Panel_bool * MDRs)
                    Case "ADA-12AWG-GS"
                        Starter_Panel(kvp.Key) = (Panel_bool * big_panels) + ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters
                    Case "ADA-W14BLKxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 5 * ((Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24"))) + (12 * Starter_Panel.Item("ADA-3VFD-MS-ADPT") + Starter_Panel.Item("ADA-3VFD-MS-1.3") + Starter_Panel.Item("ADA-3VFD-MS-2.1") _
                            + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0")) + (12 * ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters)
                    Case "ADA-W14BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 5 * ((Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24")))
                    Case "ADA-W14BRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-W14BLUxxx-xxx-x")
                    Case "ADA-W10BLKxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-2PS-40-24") * 6 + (6 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-W10BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-2PS-40-24") * 5 + (3 * (CType(Me.brakes_table(7, 0), Double))) + (3 * (CType(Me.brakes_table(13, 0), Double)))
                    Case "ADA-W10BRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-W10BLUxxx-xxx-x")
                    Case "ADA-W10ORGxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 3 * Panel_bool + (CType(Me.brakes_table(7, 0), Double) * 2) + (CType(Me.brakes_table(13, 0), Double) * 2) + (CType(Me.brakes_table(2, 0), Double) * 2) + (CType(Me.brakes_table(3, 0), Double) * 2)
                    Case "ADA-W12ORGxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 9 * Panel_bool + (CType(Me.brakes_table(7, 0), Double) * 2) + (CType(Me.brakes_table(13, 0), Double) * 2) + (CType(Me.brakes_table(2, 0), Double) * 2) + (CType(Me.brakes_table(3, 0), Double) * 2)
                    Case "ADA-W12BRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 8 * Panel_bool + (CType(Me.brakes_table(7, 0), Double) * 3) + (CType(Me.brakes_table(13, 0), Double) * 3)
                    Case "ADA-W12BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 9 * Panel_bool + (CType(Me.brakes_table(7, 0), Double) * 3) + (CType(Me.brakes_table(13, 0), Double) * 3)
                    Case "ADA-W12WHT/BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 8 * Panel_bool
                    Case "ADA-W16BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = (3 * (CType(Me.brakes_table(7, 0), Double))) + (3 * (CType(Me.brakes_table(13, 0), Double)))
                    Case "ADA-W16BRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-W16BLUxxx-xxx-x")
                    Case "ADA-W18BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 4 * Starter_Panel.Item("ADA-SWD-PFM-8")
                    Case "ADA-W18BRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-W18WHT/BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 4 * Starter_Panel.Item("ADA-SWD-PFM-8")
                    Case "ADA-W22BLUxxx-xxx-x"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-SWD-CM")
                    Case "ADA-W16WHTxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-W12GRNxxx-xxx-x"
                        Starter_Panel(kvp.Key) = 2 * (Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24")) + 3 * (Starter_Panel.Item("ADA-3VFD-MS-1.3") + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") _
                            + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0"))
                    Case "ADA-W10EYE-YEL"
                        Starter_Panel(kvp.Key) = 2 * big_panels
                    Case "ADA-W10FERRULE-YEL"
                        Starter_Panel(kvp.Key) = (12 * Starter_Panel.Item("ADA-3CB-30")) + (6 * (Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24"))) + (3 * CType(Me.brakes_table(13, 0), Double)) + (12 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-W12FERRULE-GRY"
                        Starter_Panel(kvp.Key) = 4 * Panel_bool
                    Case "ADA-W14FERRULE-BLU"
                        Starter_Panel(kvp.Key) = (6 * (Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24"))) + (12 * (Starter_Panel.Item("ADA-3VFD-MS-1.3") _
                            + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0"))) + (12 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-W16FERRULE-BLK"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-2PS-05-24") + Starter_Panel.Item("ADA-2PS-10-24") + Starter_Panel.Item("ADA-2PS-20-24") + Starter_Panel.Item("ADA-2PS-40-24")) + (Starter_Panel.Item("ADA-3VFD-MS-1.3") _
                            + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0"))
                    Case "ADA-W18FERRULE-RED"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-W22FERRULE-WHT"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-SWD-CM") * 6
                    Case "ADA-WIRE-WAY"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-WIRE-WAY-Cover"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-DIN-RAIL"
                        Starter_Panel(kvp.Key) = Panel_bool * (big_panels / 2) + ((ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters) / 2)
                    Case "ADA-WIRE-TIE"
                        Starter_Panel(kvp.Key) = Panel_bool * big_panels * 5
                    Case "ADA-WIRE-TIE-MOUNT"
                        Starter_Panel(kvp.Key) = Panel_bool * big_panels * 5

                        '------------------------- Label Parts ---------------------
                    Case "ADA-WL (WireLabel)"
                        Starter_Panel(kvp.Key) = 0.1 * big_panels * Panel_bool
                    Case "ADA-GND-L (Label)"
                        Starter_Panel(kvp.Key) = Panel_bool * Math.Round(((big_panels + (CType(Me.brakes_table(7, 0), Double)) + (CType(Me.brakes_table(13, 0), Double)) + (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))) / 100)
                    Case "ADA-Label-BLK-Rib"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-GND-L (Label)")
                    Case "ADA-Label-ORG-Arc-Flash"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-GND-L (Label)")
                    Case "ADA-Label-ORG-Arc-Flash-r"
                        Starter_Panel(kvp.Key) = 0 'Starter_Panel.Item("ADA-GND-L (Label)")
                    Case "ADA-Label-Logo"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-GND-L (Label)")
                    Case "ADA-DECAL-BLK"
                        Starter_Panel(kvp.Key) = Panel_bool * ((big_panels + (CType(Me.brakes_table(7, 0), Double)) + (CType(Me.brakes_table(13, 0), Double)) + (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters)))


                        '------------------------------- Amp Monitoring electronic motors starters---------------------------------

                    Case "ADA-3RMS-0.2-2.5"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, CType(Me.Motors_table(0, 0), Double) + CType(Me.Motors_table(12, 0), Double) + CType(Me.Motors_table(36, 0), Double), 0)
                    Case "ADA-3RMS-1.5-9.0" 'check
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (CType(Me.Motors_table(1, 0), Double) + CType(Me.Motors_table(2, 0), Double) + CType(Me.Motors_table(3, 0), Double) + CType(Me.Motors_table(4, 0), Double) + CType(Me.Motors_table(5, 0), Double) _
                            + CType(Me.Motors_table(6, 0), Double) + CType(Me.Motors_table(7, 0), Double) + CType(Me.Motors_table(8, 0), Double) + CType(Me.Motors_table(9, 0), Double)) + (CType(Me.Motors_table(13, 0), Double) + CType(Me.Motors_table(14, 0), Double) + CType(Me.Motors_table(15, 0), Double) _
                            + CType(Me.Motors_table(16, 0), Double) + CType(Me.Motors_table(17, 0), Double) + CType(Me.Motors_table(18, 0), Double) + CType(Me.Motors_table(19, 0), Double) + CType(Me.Motors_table(20, 0), Double) + CType(Me.Motors_table(21, 0), Double)) + (CType(Me.Motors_table(37, 0), Double) _
                            + CType(Me.Motors_table(38, 0), Double) + CType(Me.Motors_table(39, 0), Double) + CType(Me.Motors_table(40, 0), Double) + CType(Me.Motors_table(41, 0), Double) + CType(Me.Motors_table(42, 0), Double) + CType(Me.Motors_table(43, 0), Double) + CType(Me.Motors_table(44, 0), Double) _
                            + CType(Me.Motors_table(45, 0), Double) + (CType(Me.brakes_table(2, 0), Double)) + (CType(Me.brakes_table(13, 0), Double))), 0)
                    Case "ADA-3RMS-8.0-32.0" 'check
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (CType(Me.Motors_table(10, 0), Double) + CType(Me.Motors_table(11, 0), Double)) + (CType(Me.Motors_table(22, 0), Double) + CType(Me.Motors_table(23, 0), Double)) + (CType(Me.Motors_table(46, 0), Double) + CType(Me.Motors_table(47, 0), Double)) + (CType(Me.brakes_table(3, 0), Double)), 0)
                    Case "ADA-SWD-DIL-PKE"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, Starter_Panel.Item("ADA-3RMS-8.0-32.0"), 0)

                        '---------------------------------- Soft Start VFDs --------------------

                    Case "ADA-3VFD-MS-CBL"  'check with dave
                        Starter_Panel(kvp.Key) = If(Starter_Panel.Item("ADA-3VFD-MS-ADPT") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-1.3") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-2.1") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-3.6") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-5.0") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-8.5") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-11.0") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-16.0") > 0, 1, 0)
                    Case "ADA-3VFD-MS-STCK"  'check with dave
                        Starter_Panel(kvp.Key) = If(Starter_Panel.Item("ADA-3VFD-MS-1.3") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-2.1") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-3.6") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-5.0") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-8.5") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-11.0") > 0 Or Starter_Panel.Item("ADA-3VFD-MS-16.0") > 0, 1, 0)
                    Case "ADA-3VFD-MS-ADPT"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-3VFD-MS-1.3") + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0")
                    Case "ADA-3VFD-MS-1.3"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(0)
                    Case "ADA-3VFD-MS-2.1"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(1)
                    Case "ADA-3VFD-MS-3.6"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(2)
                    Case "ADA-3VFD-MS-5.0"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(3)
                    Case "ADA-3VFD-MS-8.5"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(4)
                    Case "ADA-3VFD-MS-11.0"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(5)
                    Case "ADA-3VFD-MS-16.0"
                        Starter_Panel(kvp.Key) = Limited_VFD_Soft_Q(6)
                    Case "ADA-3VFD-MS-ADPT"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-3VFD-MS-1.3") + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0") + Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0")

                        '---------------------------- Cable glands --------------------
                    Case "ADA-BXORING-3/4"
                        Starter_Panel(kvp.Key) = 0 '((ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters) * 6) + Panel_bool * (9 * big_panels)
                    Case "ADA-BXMHCG-0X0"
                        Starter_Panel(kvp.Key) = 0 'Starter_Panel.Item("ADA-BXORING-3/4")
                    Case "ADA-BXPCG-3/4R"
                        Starter_Panel(kvp.Key) = 0 'Starter_Panel.Item("ADA-BXORING-3/4")
                    Case "ADA-BXMHCG-5X5" 'error reference to H183 which is empty
                        Starter_Panel(kvp.Key) = 0 'If(Starter_Panel.Item("ADA-SWD-CM") > 0, 0, Math.Ceiling(1 * 16 / 5)) * Panel_bool
                    Case "ADA-BXPCN-PG9"
                        Starter_Panel(kvp.Key) = (big_panels * 2) * Panel_bool + (2 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-BXPCP-PG9"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-BXPCN-PG9")
                    Case "ADA-BXPGN-3/4"
                        Starter_Panel(kvp.Key) = 0 '(big_panels * 11) * Panel_bool + (6 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "ADA-BXPCP-3/4"
                        Starter_Panel(kvp.Key) = (big_panels * 8) * Panel_bool + (8 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))
                    Case "2M25PA"
                        Starter_Panel(kvp.Key) = 0 'big_panels * 2 * Panel_bool
                    Case "ADA-BXPCP-1"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-BXPCN-PG9")
                    Case "2M32PA"
                        Starter_Panel(kvp.Key) = 0 'big_panels * 2 * Panel_bool
                    Case "ADA-BHR5-M-M20"
                        Starter_Panel(kvp.Key) = 0 'big_panels * Panel_bool
                    Case "ADA-BXPCP-NPT1"
                        Starter_Panel(kvp.Key) = 0

                        '----------------------------- Panel ------------------------

                    Case "ADA-NEMA12-BOX30"
                        Starter_Panel(kvp.Key) = Panel_bool * ADA_NEMA12_BOX30_No_Starters
                    Case "ADA-NEMA12-BOX24"
                        Starter_Panel(kvp.Key) = Math.Round(If((big_panels * Panel_bool * NEMA12) > (Total_current_per_group / 22), big_panels * NEMA12, Math.Ceiling(Total_current_per_group / 22 * NEMA12)) + (ADA_NEMA12_BOX24_No_Starters + CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double)), MidpointRounding.AwayFromZero)
                    Case "ADA-NEMA4X-BOX24"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-SP12-21"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-PSL-MS"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")
                    Case "ADA-STRUT-MOUNT"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-MOUNT-SPACER"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-MS3/8-16/1750"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-W38"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-MS10-24/375-PHCS"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 12
                    Case "ADA-3CB-PLR"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-MS10-24/750"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-MS10-24/750-SHCS"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "1119931"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "1137341"
                        Starter_Panel(kvp.Key) = (Starter_Panel.Item("ADA-NEMA12-BOX30") + Starter_Panel.Item("ADA-NEMA12-BOX24")) * 4
                    Case "ADA-MOUNT-PAD"
                        Starter_Panel(kvp.Key) = big_panels * 4 * Panel_bool + (4 * (ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters))

                       '-------------- Misc equipment ------------------------------------------------------

                    Case "94913"
                        Starter_Panel(kvp.Key) = (big_panels * Panel_bool) + ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters + CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double)
                    Case "ADA-GTB"
                        Starter_Panel(kvp.Key) = (big_panels * Panel_bool) + ADA_NEMA12_BOX24_No_Starters + ADA_NEMA12_BOX30_No_Starters + CType(Me.brakes_table(7, 0), Double) + CType(Me.brakes_table(13, 0), Double)
                    Case "ADA-BBT1"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 0, big_panels, 0)
                    Case "ADA-BBL2"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, (Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32") + Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) / 3, 0), MidpointRounding.AwayFromZero)
                    Case "ADA-BBL3"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, (Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32") + Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) / 5, 0), MidpointRounding.AwayFromZero)
                    Case "ADA-BBL4"
                        Starter_Panel(kvp.Key) = Math.Round(If(Amp_Mon_Electro = 0, (Starter_Panel.Item("ADA-3MS-0.1-0.16") + Starter_Panel.Item("ADA-3MS-0.16-0.25") + Starter_Panel.Item("ADA-3MS-0.25-0.4") + Starter_Panel.Item("ADA-3MS- 0.4-0.63") + Starter_Panel.Item("ADA-3MS-0.63-1.0") _
                        + Starter_Panel.Item("ADA-3MS-1.0-1.6") + Starter_Panel.Item("ADA-3MS-1.6-2.5") + Starter_Panel.Item("ADA-3MS-2.5-4.0") + Starter_Panel.Item("ADA-3MS-4.0-6.3") + Starter_Panel.Item("ADA-3MS-6.3-10.0") + Starter_Panel.Item("ADA-3MS-8-12") _
                        + Starter_Panel.Item("ADA-3MS-10-16") + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32") + Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") _
                        + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") _
                        + Starter_Panel.Item("ADA-3RMS-4.0-6.3") + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25")) / 6, 0), MidpointRounding.AwayFromZero)
                    Case "ADA-BBL5"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-BBL4")
                    Case "ADA-EMS-CCL2"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0")) / 4, 0)
                    Case "ADA-EMS-CCL3"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0")) / 6, 0)
                    Case "ADA-EMS-CCL4"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0")) / 8, 0)
                    Case "ADA-EMS-CCL5"
                        Starter_Panel(kvp.Key) = If(Amp_Mon_Electro = 1, (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0")) / 10, 0)
                    Case "ADA-3FB-5"
                        Starter_Panel(kvp.Key) = (Primary_PS_5A + Primary_PS_10A + Primary_PS_20A) * Panel_bool + CType(Me.brakes_table(3, 0), Double) + (3 * CType(Me.brakes_table(7, 0), Double)) + (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0") + Starter_Panel.Item("ADA-3RMS-8.0-32.0")) / 5
                    Case "ADA-CCF-3"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-3FB-5") * 3
                    Case "ADA-CCF-10"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-CCF-30"
                        Starter_Panel(kvp.Key) = ((Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0") + Starter_Panel.Item("ADA-3RMS-8.0-32.0")) / 5) * 3
                    Case "ADA-1CB-04"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-1CB-10"
                        Starter_Panel(kvp.Key) = (CType(Me.brakes_table(3, 0), Double) * 2) * Panel_bool
                    Case "ADA-1CB-16"
                        Starter_Panel(kvp.Key) = Panel_bool * (CType(Me.bus_table(9, 0), Double))
                    Case "ADA-1CB-20"
                        Starter_Panel(kvp.Key) = Starter_Panel.Item("ADA-Hytrol MDR PWR CBL") * Panel_bool + (6 * CType(Me.brakes_table(13, 0), Double))
                    Case "ADA-1CB-30"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-1CB-BB-18"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-4CB-04"
                        Starter_Panel(kvp.Key) = 0
                    Case "ADA-8CB-04"
                        Starter_Panel(kvp.Key) = 4 * (CType(Me.brakes_table(7, 0), Double))

                End Select


            Next



            '/////////// long calculation ahead //////////////
            Panel_Space_Top = Math.Round((45 * (Math.Round(Starter_Panel.Item("ADA-3MS-0.1-0.16"), MidpointRounding.AwayFromZero) + Math.Round(Starter_Panel.Item("ADA-3MS-0.16-0.25")) + Math.Round(Starter_Panel.Item("ADA-3MS-0.25-0.4")) + Math.Round(Starter_Panel.Item("ADA-3MS- 0.4-0.63")) + Math.Round(Starter_Panel.Item("ADA-3MS-0.63-1.0")) + Math.Round(Starter_Panel.Item("ADA-3MS-1.0-1.6")) + Math.Round(Starter_Panel.Item("ADA-3MS-1.6-2.5")) _
                + Math.Round(Starter_Panel.Item("ADA-3MS-2.5-4.0")) + Math.Round(Starter_Panel.Item("ADA-3MS-4.0-6.3")) + Math.Round(Starter_Panel.Item("ADA-3MS-6.3-10.0")) + Math.Round(Starter_Panel.Item("ADA-3MS-8-12")) + Math.Round(Starter_Panel.Item("ADA-3MS-10-16")) + Starter_Panel.Item("ADA-3MS-16-20") + Starter_Panel.Item("ADA-3MS-20-25") + Starter_Panel.Item("ADA-3MS-25-32") + Starter_Panel.Item("ADA-3RMS-8.0-32.0"))) +
                (90 * (Starter_Panel.Item("ADA-3RMS-0.1-0.16") + Starter_Panel.Item("ADA-3RMS-0.16-0.25") + Starter_Panel.Item("ADA-3RMS-0.25-0.4") + Starter_Panel.Item("ADA-3RMS- 0.4-0.63") + Starter_Panel.Item("ADA-3RMS-0.63-1.0") + Starter_Panel.Item("ADA-3RMS-1.0-1.6") + Starter_Panel.Item("ADA-3RMS-1.6-2.5") + Starter_Panel.Item("ADA-3RMS-2.5-4.0") + Starter_Panel.Item("ADA-3RMS-4.0-6.3") _
                + Starter_Panel.Item("ADA-3RMS-6.3-10.0") + Starter_Panel.Item("ADA-3RMS-8-12") + Starter_Panel.Item("ADA-3RMS-10-16") + Starter_Panel.Item("ADA-3RMS-16-20") + Starter_Panel.Item("ADA-3RMS-20-25"))) _
                + (35 * (Starter_Panel.Item("ADA-SWD-PFM-8") + Starter_Panel.Item("ADA-SWD-4DI-2DO"))) + (72 * (Starter_Panel.Item("ADA-VFD-0.5-AB-575") + Starter_Panel.Item("ADA-VFD-1-AB-575") + Starter_Panel.Item("ADA-VFD-2-AB-575") + Starter_Panel.Item("ADA-VFD-3-AB-575") + Starter_Panel.Item("ADA-VFD-0.5-AB-525") + Starter_Panel.Item("ADA-VFD-1-AB-525") _
                + Starter_Panel.Item("ADA-VFD-2-AB-525") + Starter_Panel.Item("ADA-VFD-3-AB-525"))) + (87 * (Starter_Panel.Item("ADA-VFD-5-AB-575") + Starter_Panel.Item("ADA-VFD-5-AB-525"))) + (109 * (Starter_Panel.Item("ADA-VFD-7.5-AB-575") + Starter_Panel.Item("ADA-VFD-10-AB-575") + Starter_Panel.Item("ADA-VFD-7.5-AB-525") + Starter_Panel.Item("ADA-VFD-10-AB-525"))) _
                + (72 * (Starter_Panel.Item("ADA-VFD-0.5-AB-525") + Starter_Panel.Item("ADA-VFD-1-AB-525") + Starter_Panel.Item("ADA-VFD-2-AB-525") + Starter_Panel.Item("ADA-VFD-3-AB-525"))) + (87 * (Starter_Panel.Item("ADA-VFD-5-AB-525"))) + (109 * (Starter_Panel.Item("ADA-VFD-7.5-AB-525") + Starter_Panel.Item("ADA-VFD-10-AB-525"))) + (30 * (Starter_Panel.Item("ADA-3RMS-0.2-2.5") + Starter_Panel.Item("ADA-3RMS-1.5-9.0"))) _
                + (40 * Starter_Panel.Item("ADA-2PS-05-24")) + (62 * Starter_Panel.Item("ADA-2PS-10-24")) + (65 * Starter_Panel.Item("ADA-2PS-20-24")) + (110 * Starter_Panel.Item("ADA-2PS-40-24")) + (45 * (Starter_Panel.Item("ADA-3VFD-MS-1.3") + Starter_Panel.Item("ADA-3VFD-MS-2.1") + Starter_Panel.Item("ADA-3VFD-MS-3.6") + Starter_Panel.Item("ADA-3VFD-MS-5.0"))) + (90 * (Starter_Panel.Item("ADA-3VFD-MS-8.5") + Starter_Panel.Item("ADA-3VFD-MS-11.0") + Starter_Panel.Item("ADA-3VFD-MS-16.0"))) + (51 * Starter_Panel.Item("ADA-RJ45-SWT-8-port")), MidpointRounding.AwayFromZero)

            '/////////////////////////////////////////////////




            Panel_Space_Bottom = Math.Round((35 * Starter_Panel.Item("ADA-SWD4-8FRF-10")) + (22 * Starter_Panel.Item("ADA-1SRB")) + (5 * Starter_Panel.Item("ADA-1TB2_2")) + (6 * Starter_Panel.Item("ADA-1TB4")) + (8 * Starter_Panel.Item("ADA-1TB6")) + (10 * Starter_Panel.Item("ADA-1TB10")) +
                (10 * Starter_Panel.Item("ADA-1TEA")) + Math.Round(2.5 * (Starter_Panel.Item("ADA-1EB2") + Starter_Panel.Item("ADA-1EB10")), MidpointRounding.AwayFromZero) + (54 * Starter_Panel.Item("ADA-3FB-5")) + (35 * Starter_Panel.Item("ADA-4CB-04")) + (70 * Starter_Panel.Item("ADA-8CB-04")), MidpointRounding.AwayFromZero)



            big_panels = If((Panel_Space_Top / panel_width) > (Panel_Space_Bottom / panel_width), Math.Ceiling(Panel_Space_Top / panel_width), Math.Ceiling(Panel_Space_Bottom / panel_width))

            'total current (H318)
            total_current = Math.Round((non_reversing_amps + reversing_amps + VFD_amps + Soft_start_amps) * If(big_panels > 0, 1, 0))
            Total_current_per_group = total_current + (1.02 * CType(Me.brakes_table(2, 0), Double)) + (4.16 * CType(Me.brakes_table(3, 0), Double)) + (CType(Me.brakes_table(0, 0), Double) / 4) +
                   (CType(Me.brakes_table(12, 0), Double) * 0.0331) + (Starter_Panel.Item("ADA-2PS-05-24") * 0.6) + (Starter_Panel.Item("ADA-2PS-10-24") * 0.6) + (Starter_Panel.Item("ADA-2PS-20-24") * 0.65) + (Starter_Panel.Item("ADA-2PS-40-24") * 1.35)


            '-------------------------------------------------------
        Next


        Total_swire_nodes = Starter_Panel.Item("ADA-SWD-CM") + Starter_Panel.Item("ADA-SWD-PFM-8") + Starter_Panel.Item("ADA-SWD-4DI-2DO")
        busbar_slots = (2 * reversing_starters) + non_reversing_starters + VFD_starters + Soft_starters

    End Sub



    'This will calculate the ADA parts in the IO Panel (IO Panel tab)
    Public Sub Calculate_IO_Panel()

        '///////////////////////////// Setup ////////////////////////////////

        '--------- Inputs calculation -----------------
        Inputs = 0

        For i = 0 To 47
            Inputs = Inputs + (CType(Me.Motors_table(i, 1), Double) * CType(Me.Motors_table(i, 0), Double))
        Next

        For i = 0 To 18
            Inputs = Inputs + (CType(Me.inputs_table(i, 1), Double) * CType(Me.inputs_table(i, 0), Double))
        Next

        For i = 0 To 18
            Inputs = Inputs + (CType(Me.push_table(i, 1), Double) * CType(Me.push_table(i, 0), Double))
        Next

        For i = 0 To 14
            Inputs = Inputs + (CType(Me.lights_table(i, 1), Double) * CType(Me.lights_table(i, 0), Double))
        Next

        For i = 0 To 11
            Inputs = Inputs + (CType(Me.field_table(i, 1), Double) * CType(Me.field_table(i, 0), Double))
        Next
        For i = 0 To 7  'check
            Inputs = Inputs + (CType(Me.bus_table(i, 1), Double) * CType(Me.bus_table(i, 0), Double))
        Next
        For i = 0 To 3
            Inputs = Inputs + (CType(Me.scanner_table(i, 1), Double) * CType(Me.scanner_table(i, 0), Double))
        Next
        For i = 0 To 14
            Inputs = Inputs + (CType(Me.brakes_table(i, 1), Double) * CType(Me.brakes_table(i, 0), Double))
        Next


        Inputs = Math.Abs(Inputs + Math.Ceiling((CType(Me.Percent_overage_inputs, Double) / 100) * (Inputs - 8)))

        Dim bus_in As Double : bus_in = 0
        For i = 0 To 7  'check
            bus_in = bus_in + (CType(Me.bus_table(i, 1), Double) * CType(Me.bus_table(i, 0), Double))
        Next

        cr_input = If((Inputs - bus_in - 8) * Panel_bool > 0, (Inputs - bus_in - 8) * Panel_bool, 0)


        '-------------------------------------------------------------------------------------------------

        Monitoring_Inputs = 0  'now IO link

        '---------------------------------------------------------------------
        'Non Motion Outputs Calculations

        Non_Motion_Outputs = 0


        For i = 0 To 47
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.Motors_table(i, 2), Double) * CType(Me.Motors_table(i, 0), Double))
        Next

        For i = 0 To 18
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.inputs_table(i, 2), Double) * CType(Me.inputs_table(i, 0), Double))
        Next

        For i = 0 To 18
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.push_table(i, 2), Double) * CType(Me.push_table(i, 0), Double))
        Next

        For i = 0 To 14
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.lights_table(i, 2), Double) * CType(Me.lights_table(i, 0), Double))
        Next

        For i = 0 To 11
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.field_table(i, 2), Double) * CType(Me.field_table(i, 0), Double))
        Next
        For i = 0 To 7  'check
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.bus_table(i, 2), Double) * CType(Me.bus_table(i, 0), Double))
        Next
        For i = 0 To 3
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.scanner_table(i, 2), Double) * CType(Me.scanner_table(i, 0), Double))
        Next
        For i = 0 To 14
            Non_Motion_Outputs = Non_Motion_Outputs + (CType(Me.brakes_table(i, 2), Double) * CType(Me.brakes_table(i, 0), Double))
        Next


        Non_Motion_Outputs = Math.Abs(Math.Ceiling(Non_Motion_Outputs * ((CType(Me.Percent_overage_nm_outputs, Double) / 100) + 1)))

        '-----cr

        Dim bus_nm As Double : bus_nm = 0
        For i = 0 To 7  'check
            bus_nm = bus_nm + (CType(Me.bus_table(i, 1), Double) * CType(Me.bus_table(i, 0), Double))
        Next

        cr_n_motion = If((Non_Motion_Outputs - bus_nm) * Panel_bool > 0, (Non_Motion_Outputs - bus_nm) * Panel_bool, 0)



        '--------------------------------------------------------------
        motion_outputs = 0

        For i = 0 To 47
            motion_outputs = motion_outputs + (CType(Me.Motors_table(i, 3), Double) * CType(Me.Motors_table(i, 0), Double))
        Next

        For i = 0 To 18
            motion_outputs = motion_outputs + (CType(Me.inputs_table(i, 3), Double) * CType(Me.inputs_table(i, 0), Double))
        Next

        For i = 0 To 18
            motion_outputs = motion_outputs + (CType(Me.push_table(i, 3), Double) * CType(Me.push_table(i, 0), Double))
        Next

        For i = 0 To 14
            motion_outputs = motion_outputs + (CType(Me.lights_table(i, 3), Double) * CType(Me.lights_table(i, 0), Double))
        Next

        For i = 0 To 11
            motion_outputs = motion_outputs + (CType(Me.field_table(i, 3), Double) * CType(Me.field_table(i, 0), Double))
        Next
        For i = 0 To 11
            motion_outputs = motion_outputs + (CType(Me.bus_table(i, 3), Double) * CType(Me.bus_table(i, 0), Double))
        Next
        For i = 0 To 3
            motion_outputs = motion_outputs + (CType(Me.scanner_table(i, 3), Double) * CType(Me.scanner_table(i, 0), Double))
        Next
        For i = 0 To 14
            motion_outputs = motion_outputs + (CType(Me.brakes_table(i, 3), Double) * CType(Me.brakes_table(i, 0), Double))
        Next


        motion_outputs = Math.Abs(Math.Ceiling(motion_outputs * ((CType(Me.Percent_overage_m_outputs, Double) / 100) + 1)))

        Dim bus_m As Double : bus_m = 0
        For i = 0 To 7  'check
            bus_m = bus_m + (CType(Me.bus_table(i, 1), Double) * CType(Me.bus_table(i, 0), Double))
        Next

        cr_motion = If((motion_outputs - bus_m) * Panel_bool > 0, (motion_outputs - bus_m) * Panel_bool, 0)

        '---------------------------------------------------------------
        Dim sum_motors As Double : sum_motors = 0

        For i = 0 To 45
            sum_motors = sum_motors + (CType(Me.Motors_table(i, 0), Double))
        Next

        has_wire = If((Panel_bool * (sum_motors) + (CType(Me.bus_table(2, 0), Double) + CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double) _
            + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double))) > 0, 1, 0)

        system_power_req = If((Inputs > 0 Or Non_Motion_Outputs > 0 Or motion_outputs > 0), 1, 0)


        '///////////////////////////      I/O 4 Point Cards         /////////////////////////////////
        'IN

        IO_4_inputs = 0
        IO_4_non_motion = If(((Non_Motion_Outputs Mod 16 <= 4) And (Non_Motion_Outputs Mod 16 > 0) And (CType(Use_4pts_IO, Double) = 1)) = True, 1, 0)
        IO_4_motion = If(((motion_outputs Mod 16 <= 4) And (motion_outputs Mod 16 > 0) And (CType(Use_4pts_IO, Double) = 1)) = True, 1, 0)

        Totals_IO_4 = If(Remote_IO > 0, 0, (Panel_bool * IO_4_inputs))
        IO_4_convert_4 = If(((Totals_IO_4 Mod 4 <= 1) And (Totals_IO_4 Mod 4 > 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1) And (CType(Use_4pts_IO_RB, Double) = 1)) = True, 1, 0)
        IO_4_convert_8 = If((((Totals_IO_4 - IO_4_convert_4) Mod 4) <= 2 And ((Totals_IO_4 - IO_4_convert_4) Mod 4) > 0 And (CType(Use_8pts_IO_RB, Double) = 1)) = True, 1, 0)
        IO_4_ADA_SWD_FF_Coupling = If(Math.Ceiling(Totals_IO_4 - IO_4_convert_4 - IO_4_convert_8 * 2) / 4 > 0, Math.Ceiling(Totals_IO_4 - IO_4_convert_4 - IO_4_convert_8 * 2) / 4, 0)
        IO_4_Pre_combine_8s = IO_4_convert_4
        IO_4_total_JB = IO_4_Pre_combine_8s

        IO_4_pigtail_single = If(((CType(M23_at_IO_Receptacle, Double) = 0) And (CType(M23_Bulkhead_at_ADA, Double) = 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1)) = True, IO_4_total_JB * CType(Use_4pts_IO_RB, Double), 0)
        IO_4_conn_At_JB_single = If((((CType(M23_at_IO_Receptacle, Double) = 1) Or (CType(M23_Bulkhead_at_ADA, Double) = 1)) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1)) = True, IO_4_total_JB * CType(Use_4pts_IO_RB, Double), 0)

        '------ needs checking 100 value should be N63 avoid redirection
        IO_4_pigtail_double = If(((CType(M23_at_IO_Receptacle, Double) = 0) And (CType(M23_Bulkhead_at_ADA, Double) = 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 0)) = True, 100 * CType(Use_4pts_IO_RB, Double), 0)
        IO_4_conn_At_JB_double = If((((CType(M23_at_IO_Receptacle, Double) = 1) Or (CType(M23_Bulkhead_at_ADA, Double) = 1)) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 0)) = True, 100 * CType(Use_4pts_IO_RB, Double), 0)

        'OUT
        Totals_IO_4_out = If(Remote_IO > 0, 0, (Panel_bool * (IO_4_non_motion + IO_4_motion)))
        IO_4_convert_4_out = If(((Totals_IO_4_out Mod 4 <= 1) And (Totals_IO_4_out Mod 4 > 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1) And (CType(Use_4pts_IO_RB, Double) = 1)) = True, 1, 0)
        IO_4_convert_8_out = If((((Totals_IO_4_out - IO_4_convert_4_out) Mod 4) <= 2 And ((Totals_IO_4_out - IO_4_convert_4_out) Mod 4) > 0 And (CType(Use_8pts_IO_RB, Double) = 1)) = True, 1, 0)
        IO_4_convert_16_out = If(Math.Ceiling(Totals_IO_4_out - IO_4_convert_4_out - IO_4_convert_8_out * 2) / 4 > 0, Math.Ceiling(Totals_IO_4_out - IO_4_convert_4_out - IO_4_convert_8_out * 2) / 4, 0)
        IO_4_Pre_combine_8s_out = IO_4_convert_4_out
        IO_4_total_JB_out = IO_4_Pre_combine_8s_out

        IO_4_pigtail_single_out = If(((CType(M23_at_IO_Receptacle, Double) = 0) And (CType(M23_Bulkhead_at_ADA, Double) = 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1)) = True, IO_4_total_JB_out * CType(Use_4pts_IO_RB, Double), 0)
        IO_4_conn_At_JB_single_out = If((((CType(M23_at_IO_Receptacle, Double) = 1) Or (CType(M23_Bulkhead_at_ADA, Double) = 1)) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 1)) = True, IO_4_total_JB_out * CType(Use_4pts_IO_RB, Double), 0)


        '///////////////////////////      I/O 8 Point Cards         /////////////////////////////////

        'IN
        IO_8_inputs = If(((((Inputs + Monitoring_Inputs) - (IO_4_inputs * 4)) Mod 16 <= 8 And ((Inputs + Monitoring_Inputs) - (IO_4_inputs * 4)) Mod 16 > 0 And Use_8pts_IO = 1) Or (Use_16pts_IO And (Inputs + Monitoring_Inputs > 4))), If((Use_16pts_IO = 0 And Inputs + Monitoring_Inputs > 4), Math.Ceiling((Inputs + Monitoring_Inputs) / 8), 0), 0)
        IO_8_non_motion = If((((Non_Motion_Outputs - (IO_4_non_motion * 4)) Mod 16 <= 8) And ((Non_Motion_Outputs - (IO_4_non_motion * 4)) Mod 16 > 0) And (CType(Use_8pts_IO, Double) = 1)) Or (CType(Use_16pts_IO, Double) = 0 And Non_Motion_Outputs > 4), If((CType(Use_16pts_IO, Double) = 0 And Inputs > 4), Math.Ceiling(Non_Motion_Outputs / 8), 0), 0)
        IO_8_motion = If((((motion_outputs - (IO_4_motion * 4)) Mod 16 <= 8) And ((motion_outputs - (IO_4_motion * 4)) Mod 16 > 0) And (CType(Use_8pts_IO, Double) = 1)) Or (CType(Use_16pts_IO, Double) = 0 And motion_outputs > 4), If((CType(Use_16pts_IO, Double) = 0 And Inputs > 4), Math.Ceiling(Non_Motion_Outputs / 8), 0), 0)

        Totals_IO_8 = If(Remote_IO > 0, 0, (Panel_bool * IO_8_inputs))
        IO_8_convert_4 = 0
        IO_8_convert_8 = If((((Totals_IO_8) Mod 2) <= 1 And ((Totals_IO_8) Mod 2) > 0 And (CType(Use_8pts_IO_RB, Double) = 1)) = True, 1, 0)
        IO_8_convert_16 = If(Math.Ceiling(Totals_IO_8 - IO_8_convert_8) / 2 > 0, Math.Ceiling(Totals_IO_8 - IO_8_convert_8) / 2, 0)
        IO_8_Pre_combine_8s = IO_8_convert_8 + IO_4_convert_8
        IO_8_total_JB = IO_8_Pre_combine_8s Mod 2

        IO_8_pigtail_single = If((CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, IO_8_total_JB, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, (Math.Ceiling(Inputs / 8) + IO_8_total_JB), 0)), 0)
        IO_8_conn_At_JB_single = If((CType(M23_at_IO_Receptacle, Double) = 1 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, IO_8_total_JB, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, (Math.Ceiling(Inputs / 8) + IO_8_total_JB), 0)), 0)
        IO_8_pigtail_double = If((CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, IO_8_total_JB, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, (Math.Ceiling(Inputs / 16) + IO_8_total_JB), 0)), 0)
        IO_8_conn_At_JB_double = If((CType(M23_at_IO_Receptacle, Double) = 1 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, IO_8_total_JB, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, (Math.Ceiling(Inputs / 16) + IO_8_total_JB), 0)), 0)

        'OUT
        Totals_IO_8_out = If(Remote_IO > 0, 0, (Panel_bool * (IO_8_non_motion + IO_8_motion)))
        IO_8_convert_4_out = 0
        IO_8_convert_8_out = If((((Totals_IO_8_out) Mod 2) <= 1 And ((Totals_IO_8_out) Mod 2) > 0 And (CType(IO_4_convert_4, Double) = 1)) = True, 1, 0)
        IO_8_convert_16_out = If(Math.Ceiling(Totals_IO_8_out - IO_8_convert_8_out) / 2 > 0, Math.Ceiling(Totals_IO_8_out - IO_8_convert_8_out) / 2, 0)
        IO_8_Pre_combine_8s_out = If(IO_4_non_motion, IO_4_convert_8_out + IO_8_convert_8_out, 0)
        IO_8_total_JB_out = IO_8_Pre_combine_8s_out Mod 2

        IO_8_pigtail_single_out = If((CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, IO_8_total_JB_out, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, (Math.Ceiling((Non_Motion_Outputs + motion_outputs) / 8) + IO_8_total_JB_out), 0)), 0)
        IO_8_conn_At_JB_single_out = If((CType(M23_at_IO_Receptacle, Double) = 1 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, IO_8_total_JB_out, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, (Math.Ceiling((Non_Motion_Outputs + motion_outputs) / 8) + IO_8_total_JB_out), 0)), 0)
        IO_8_pigtail_double_out = If((CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, IO_8_total_JB_out, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, (Math.Ceiling((Non_Motion_Outputs + motion_outputs) / 16) + IO_8_total_JB_out), 0)), 0)
        IO_8_conn_At_JB_double_out = If((CType(M23_at_IO_Receptacle, Double) = 1 And CType(M23_Bulkhead_at_ADA, Double) = 0), If(CType(Use_16pts_IO_RB_Inputs, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1, IO_8_total_JB_out, If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 0, (Math.Ceiling((Non_Motion_Outputs + motion_outputs) / 16) + IO_8_total_JB_out), 0)), 0)

        '///////////////////////////      I/O 16 Point Cards         /////////////////////////////////

        '--------------------- I/O 16 Point Cards

        'IN
        IO_16_inputs = If(Math.Round(((Inputs + Monitoring_Inputs) - (IO_8_inputs * 8) + (IO_4_inputs * 4)) / 16, MidpointRounding.AwayFromZero) > 0 And CType(Use_16pts_IO, Double) = 1, Math.Ceiling(((Inputs + Monitoring_Inputs) - (IO_8_inputs * 8) + (IO_4_inputs * 4)) / 16), 0)
        IO_16_non_motion = If(Math.Round((Non_Motion_Outputs - (IO_8_non_motion * 8) + (IO_4_non_motion * 4)) / 16, MidpointRounding.AwayFromZero) > 0 And CType(Use_16pts_IO, Double) = 1, Math.Ceiling((Non_Motion_Outputs - (IO_8_non_motion * 8) + (IO_4_non_motion * 4)) / 16), 0)
        IO_16_motion = If(Math.Round((motion_outputs - (IO_8_motion * 8) + (IO_4_motion * 4)) / 16, MidpointRounding.AwayFromZero) > 0 And CType(Use_16pts_IO, Double) = 1, Math.Ceiling((motion_outputs - (IO_8_motion * 8) + (IO_4_motion * 4)) / 16), 0)

        Totals_IO_16 = If(Math.Ceiling((Inputs - (IO_8_inputs * 8) + (IO_4_inputs * 4)) / 16) > 0 And CType(Use_16pts_IO, Double) = 1, Math.Ceiling((Inputs - (IO_8_inputs * 8) + (IO_4_inputs * 4)) / 16), 0)
        IO_16_convert_4 = 0
        IO_16_convert_8 = 0
        IO_16_convert_16 = Totals_IO_16
        IO_16_Pre_combine_8s = IO_4_ADA_SWD_FF_Coupling + IO_8_convert_16 + IO_16_convert_16
        IO_16_total_JB = IO_16_Pre_combine_8s + Math.Floor(IO_8_Pre_combine_8s / 2)

        IO_16_pigtail_single = If(CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1 And CType(Use_16pts_IO_RB_Inputs, Double) = 1, IO_16_total_JB, 0)
        ' weird OR(1,n11=1,n12=1) is always 1
        IO_16_conn_At_JB_single = If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1 And CType(Use_16pts_IO_RB_Inputs, Double) = 1, IO_16_total_JB, 0)
        IO_16_pigtail_double = 0
        IO_16_conn_At_JB_double = 0

        'OUT
        Totals_IO_16_out = If(Remote_IO > 0, 0, (Panel_bool * (IO_16_non_motion + IO_16_motion)))
        IO_16_convert_4_out = 0
        IO_16_convert_8_out = 0
        IO_16_convert_16_out = Totals_IO_16_out
        IO_16_Pre_combine_8s_out = IO_4_convert_16_out + IO_8_convert_16_out + IO_16_convert_16_out
        IO_16_total_JB_out = IO_16_Pre_combine_8s_out + Math.Floor(IO_8_Pre_combine_8s_out / 2)

        IO_16_pigtail_single_out = If(CType(M23_at_IO_Receptacle, Double) = 0 And CType(M23_Bulkhead_at_ADA, Double) = 0 And CType(Single_Channel_IO_1_Input_per_RB, Double) = 1 And CType(Use_16pts_IO_RB_Inputs, Double) = 1, IO_16_total_JB_out, 0)
        IO_16_conn_At_JB_single_out = If(CType(Single_Channel_IO_1_Input_per_RB, Double) = 1 And CType(Use_16pts_IO_RB_Inputs, Double) = 1, IO_16_total_JB_out, 0)
        IO_16_pigtail_double_out = 0
        IO_16_conn_At_JB_double_out = 0

        '-------------------------------------------------------------
        '------ needs checking 100 value should be N63 avoid redirection (extra IO4)
        IO_4_pigtail_double_out = If(((CType(M23_at_IO_Receptacle, Double) = 0) And (CType(M23_Bulkhead_at_ADA, Double) = 0) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 0)) = True, IO_8_total_JB_out * CType(Use_4pts_IO_RB, Double), 0)
        IO_4_conn_At_JB_double_out = If((((CType(M23_at_IO_Receptacle, Double) = 1) Or (CType(M23_Bulkhead_at_ADA, Double) = 1)) And (CType(Single_Channel_IO_1_Input_per_RB, Double) = 0)) = True, IO_8_total_JB_out * CType(Use_4pts_IO_RB, Double), 0)


        For i = 1 To 5
            For Each kvp As KeyValuePair(Of String, Double) In IO_Panel.ToArray

                '----------- Misc always used parts -----------------------
                Select Case kvp.Key
                    Case "ADA-IOM-EGW-EIP"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, If(IO_Panel.Item("ADA-IOM-PGW-EIP") > 0, 0, 1) * Panel_bool) * removal_msbox
                    Case "ADA-IOM-EGW-EIPB"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-EGW-EIP") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-EGW-EIPF"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-EGW-EIP") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-EGW-EIPT"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-EGW-EIPF") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-PGW-EIP"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (Panel_bool * PLC_GateWay_PGW)) * removal_msbox
                    Case "ADA-EN-PB-NO-B"
                        IO_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-EN-PB-LIT"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-EN-PB-Legend"
                        IO_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-DIN-RAIL-Tall"
                        IO_Panel(kvp.Key) = Panel_bool / 2
                    Case "ADA-6x10-W-TLBL"
                        IO_Panel(kvp.Key) = Panel_bool / 100
                    Case "ADA-1EB2"
                        IO_Panel(kvp.Key) = Panel_bool
                    Case "ADA-4CB-04"
                        IO_Panel(kvp.Key) = If((IO_Panel.Item("ADA-IOM-O04M") + IO_Panel.Item("ADA-IOM-O04-Base") + IO_Panel.Item("ADA-IOM-O08E") + IO_Panel.Item("ADA-IOM-O16E")) > 3, 0, 1) * Panel_bool
                    Case "ADA-8CB-04"
                        IO_Panel(kvp.Key) = If((IO_Panel.Item("ADA-IOM-O04M") + IO_Panel.Item("ADA-IOM-O04-Base") + IO_Panel.Item("ADA-IOM-O08E") + IO_Panel.Item("ADA-IOM-O16E")) > 3, 1, 0) * Panel_bool


                        '------------------------ safety parts ------------
                    Case "ADA-IOM-PFM"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (If(motion_outputs > 0, 1, 0) * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-PFB"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-PFM") * Panel_bool)) * removal_msbox
                    Case "ADA-1SRB"
                        IO_Panel(kvp.Key) = (removal_msbox * Panel_bool + Panel_bool) + M23_at_IO_Receptacle * Panel_bool
                    Case "ADA-1SRY"
                        IO_Panel(kvp.Key) = ((removal_msbox * Panel_bool + Panel_bool) + M23_at_IO_Receptacle * Panel_bool) * Panel_bool

                        '--------------------------I/0------------------------
                    Case "ADA-IOM-I04M"
                        'Input (N22) value is always zero
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (Panel_bool * 0)) * removal_msbox
                    Case "ADA-IOM-I04B"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-I04M") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-I08E"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (Panel_bool * IO_8_inputs)) * removal_msbox
                    Case "ADA-IOM-I16E"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, ((If(IO_16_inputs < 0, 1, IO_16_inputs)) * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-O16EB"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-I16E") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-O16ET"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-I16E") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-O04M"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, ((IO_4_inputs + IO_4_non_motion) * 2)) * removal_msbox
                    Case "ADA-IOM-O04-Base"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-O04M") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-O08E"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, ((IO_8_non_motion + IO_8_motion) * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-O16E"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, ((IO_16_non_motion + IO_16_motion) * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-SM"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, ((CType(Me.scanner_table(1, 0), Double)) * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-SB"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-SM") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-CTM"
                        IO_Panel(kvp.Key) = If(Remote_IO = 1, 0, (((CType(Me.field_table(9, 0), Double) + (CType(Me.field_table(10, 0), Double)) * Panel_bool) / 2))) * removal_msbox
                    Case "ADA-IOM-CTB"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-CTM") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-CT-Term"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_Panel.Item("ADA-IOM-CTM") * Panel_bool)) * removal_msbox
                    Case "ADA-IOM-ANA-IN"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (CType(Me.field_table(11, 0), Double) * Panel_bool)) * removal_msbox

                        '-------------------- B&R X2X ---------------
                    Case "ADA-IOM-X2XModule"
                        IO_Panel(kvp.Key) = CType(Me.bus_table(0, 0), Double) * removal_msbox
                    Case "ADA-IOM-X2XBase"
                        IO_Panel(kvp.Key) = CType(Me.bus_table(0, 0), Double) * Panel_bool * removal_msbox
                    Case "ADA-IOM-12-X2XTerm"
                        IO_Panel(kvp.Key) = CType(Me.bus_table(0, 0), Double) * Panel_bool * removal_msbox

                        '----------------------SWIRE --------------
                    Case "ADA-IOM-SWD"
                        IO_Panel(kvp.Key) = (Panel_bool * If(Me.Total_swire_nodes > 98, 2, 1) + (Math.Ceiling((CType(Me.bus_table(2, 0), Double) + CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 98) * If(has_wire > 0, 1, 0) * Panel_bool)) * removal_msbox
                    Case "ADA SWD-8DI"
                        IO_Panel(kvp.Key) = Remote_IO * Panel_bool * removal_msbox

                        '--------------------------IO Link Master-----------------
                    Case "ADA-IOM-IOLINK"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, CType(Me.bus_table(8, 0), Double)) * removal_msbox
                    Case "ADA-IOM-IOLINK-Base"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, CType(Me.bus_table(8, 0), Double)) * Panel_bool * removal_msbox
                    Case "ADA-IOM-IOLINK-Term"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, CType(Me.bus_table(8, 0), Double)) * Panel_bool * removal_msbox

                        '-------------------------- Standard Terminal & Power distribution ---------------
                    Case "ADA-1TB2_2"
                        IO_Panel(kvp.Key) = If(removal_msbox > 0, 10, 4)
                    Case "ADA-1TB2X2"
                        IO_Panel(kvp.Key) = If(removal_msbox > 0, 4, 2)
                    Case "ADA-1TBG"
                        IO_Panel(kvp.Key) = 3 * Panel_bool
                    Case "ADA-1TEA"
                        IO_Panel(kvp.Key) = Panel_bool
                    Case "ADA-1TB2-JP10"
                        IO_Panel(kvp.Key) = 2 * Panel_bool * system_power_req * removal_msbox
                    Case "ADA-5x9-W-TLBL"
                        IO_Panel(kvp.Key) = 4 * Panel_bool + 10 * Panel_bool * system_power_req

                        '--------------------------------- M12 Connectors -------------------
                    Case "ADA-BHR5-M-PG-9"
                        IO_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-BHR5-F-PG-9"
                        IO_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-PBC4-03-G"
                        IO_Panel(kvp.Key) = 2 * M23_at_IO_Receptacle * Panel_bool
                    Case "ADA-SAC5-05-Y"
                        IO_Panel(kvp.Key) = M23_at_IO_Receptacle * Panel_bool
                    Case "ADA-SAC5-05-G-PWR"
                        IO_Panel(kvp.Key) = M23_at_IO_Receptacle * Panel_bool

                        '--------------------------------- Control Receptacles -------------------
                    Case "ADA-SARB081-5m-LED"
                        IO_Panel(kvp.Key) = If(Remote_IO < 1, ((IO_8_pigtail_single + IO_8_pigtail_single_out) * Panel_bool + Panel_bool), 0) * removal_msbox
                    Case "ADA-SARB082-5m-LED"
                        IO_Panel(kvp.Key) = If(Remote_IO < 1, (IO_8_pigtail_double + IO_8_pigtail_double_out) * Panel_bool, 0) * removal_msbox

                        '-------------------------------- 12 Pin ------------------------
                    Case "ADA-SARB081-m23P12"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_8_conn_At_JB_single + IO_8_pigtail_double_out) * Panel_bool) * removal_msbox

                        '---------------------------------19 pin ---------------------
                    Case "ADA-SARB042-m23P19-LED"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_4_conn_At_JB_double + IO_4_conn_At_JB_double_out) * Panel_bool) * removal_msbox
                    Case "ADA-SARB082-m23P19-LED"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_8_conn_At_JB_double + IO_8_conn_At_JB_double_out) * Panel_bool) * removal_msbox
                    Case "ADA-SARB161-m23P19-LED"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, (IO_16_pigtail_single + IO_16_conn_At_JB_single + IO_16_pigtail_single_out + IO_16_conn_At_JB_single_out) * Panel_bool) * removal_msbox

                        '----------------------------------12 pin --------------------

                    Case "ADA-SABH08112S-Ext"   'problem with the formula =IF(Setup!F11>0,0,(IF(N12=1,#REF!+N181,0)*N$2))
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-SARC-m23S12-05-P12  (Sockets to Pins)" 'depends on the part above
                        IO_Panel(kvp.Key) = 0

                        '---------------------------------19 pin --------------------
                    Case "ADA-SARC19-05-S"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, If(CType(M23_Bulkhead_at_ADA, Double) = 0, IO_Panel.Item("ADA-SARB042-m23P19-LED") + IO_Panel.Item("ADA-SARB082-m23P19-LED") + IO_Panel.Item("ADA-SARB161-m23P19-LED"), 0) * Panel_bool) * removal_msbox
                    Case "ADA-SABH08219S-Ext"
                        IO_Panel(kvp.Key) = If(Remote_IO > 0, 0, If(CType(M23_Bulkhead_at_ADA, Double) = 1, IO_Panel.Item("ADA-SARB042-m23P19-LED") + IO_Panel.Item("ADA-SARB082-m23P19-LED") + IO_Panel.Item("ADA-SARB161-m23P19-LED"), 0) * Panel_bool) * removal_msbox
                    Case "ADA-SARC-m23S19-05-P19  (Sockets to Pins)"
                        IO_Panel(kvp.Key) = Panel_bool * IO_Panel.Item("ADA-SABH08219S-Ext") * removal_msbox
                    Case "ADA-SARC021-m12-1M-MFF"
                        IO_Panel(kvp.Key) = (Math.Ceiling((IO_Panel.Item("ADA-SARB082-5m-LED") + IO_Panel.Item("ADA-SARB042-m23P19-LED") + IO_Panel.Item("ADA-SARB082-m23P19-LED")) * 8 * (Splitter_Perc_dual_ch / 100)) + (Panel_bool * CType(Me.inputs_table(13, 0), Double))) * Panel_bool * removal_msbox

                        '---------------------- Cable glands --------------

                    Case "ADA-BXPCG-3/4"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-BXPCG-3/4R"
                        IO_Panel(kvp.Key) = (IO_Panel.Item("ADA-SARB081-5m-LED") + IO_Panel.Item("ADA-SARB082-5m-LED")) * Panel_bool * removal_msbox + (Panel_bool * M23_at_IO_Receptacle)
                    Case "ADA-BXMHCG-4X6.5"
                        IO_Panel(kvp.Key) = Panel_bool * M23_at_IO_Receptacle
                    Case "ADA-BXMHCG-5X5"
                        IO_Panel(kvp.Key) = 0 'Panel_bool
                    Case "ADA-BXMHCG-0X0"  'there is a MAX excel formula. I don't know why
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-BXPGN-3/4"
                        IO_Panel(kvp.Key) = Panel_bool * (IO_Panel.Item("ADA-BXPCG-3/4") + IO_Panel.Item("ADA-BXPCG-3/4R")) * removal_msbox + (Panel_bool * M23_at_IO_Receptacle)
                    Case "ADA-BXORING-3/4"
                        IO_Panel(kvp.Key) = (IO_Panel.Item("ADA-SARB081-5m-LED") + IO_Panel.Item("ADA-SARB082-5m-LED")) * Panel_bool + (Panel_bool * M23_at_IO_Receptacle)
                    Case "2M25PA"
                        IO_Panel(kvp.Key) = 0 '3 * Panel_bool
                    Case "ADA-BXPCP-3/4"
                        IO_Panel(kvp.Key) = Panel_bool * (8 * (If(NEMA12, 1, 0) * Panel_bool + If(NEMA4X, 1, 0) * Panel_bool)) * removal_msbox
                    Case "2M32PA"
                        IO_Panel(kvp.Key) = 0 '3 * Panel_bool
                    Case "ADA-BXPCP-1"
                        IO_Panel(kvp.Key) = Panel_bool * (2 * (If(NEMA12, 1, 0) * Panel_bool + If(NEMA4X, 1, 0) * Panel_bool)) * removal_msbox

                        '------------------ Panel -------------------------
                    Case "ADA-NEMA12-BOX24"
                        IO_Panel(kvp.Key) = If(NEMA12, 1, 0) * Panel_bool * removal_msbox
                    Case "ADA-NEMA4X-BOX24"
                        IO_Panel(kvp.Key) = If(NEMA4X, 1, 0) * Panel_bool * removal_msbox
                    Case "ADA-SP12-21"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-100/100BX-G"
                        IO_Panel(kvp.Key) = Panel_bool * M23_at_IO_Receptacle
                    Case "ADA-100/100BX-TP"
                        IO_Panel(kvp.Key) = Panel_bool * M23_at_IO_Receptacle
                    Case "ADA-150/150BX-G"
                        IO_Panel(kvp.Key) = Panel_bool * Single_Channel_IO_1_Input_per_RB
                    Case "ADA-150/150BX-TP"
                        IO_Panel(kvp.Key) = Panel_bool * Use_4pts_IO_RB
                    Case "ADA-PSL-IO"
                        IO_Panel(kvp.Key) = Panel_bool * If(removal_msbox > 0, 1, 0.25) '(IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24"))
                    Case "ADA-STRUT-MOUNT"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-MOUNT-SPACER"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-MS3/8-16/1750"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-W38"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-MS10-24/375-PHCS"
                        IO_Panel(kvp.Key) = Panel_bool * 12 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-3CB-PLR"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-MS10-24/750"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-MS10-24/750-SHCS"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "1119931"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "1137341"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-MOUNT-PAD"
                        IO_Panel(kvp.Key) = Panel_bool * 4 * (IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * removal_msbox
                    Case "ADA-Label-BLK-Rib"
                        IO_Panel(kvp.Key) = 0 '(IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 200
                    Case "ADA-Label-ORG-Arc-Flash"
                        IO_Panel(kvp.Key) = 0 '(IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 200
                    Case "ADA-Label-ORG-Arc-Flash-r"
                        IO_Panel(kvp.Key) = 0 '(IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 200
                    Case "ADA-Label-Logo"
                        IO_Panel(kvp.Key) = 0 '(IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 200

                        '------------------------------ Ethernet Parts -----------------------
                    Case "ADA-RJ45-SWT-8-port"
                        IO_Panel(kvp.Key) = 0

                        '------------------------- Wiring -------------------

                    Case "ADA_I/O Harness"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-12AWG-GS"
                        IO_Panel(kvp.Key) = Panel_bool * removal_msbox
                    Case "ADA-W10BLKxxx-xxx-x"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-W10BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = 6 * Panel_bool * removal_msbox
                    Case "ADA-W10BRNxxx-xxx-x"
                        IO_Panel(kvp.Key) = 6 * Panel_bool * removal_msbox
                    Case "ADA-W12ORGxxx-xxx-x"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-W12BRNxxx-xxx-x"
                        IO_Panel(kvp.Key) = 6 * Panel_bool * removal_msbox
                    Case "ADA-W12BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = 6 * Panel_bool * removal_msbox
                    Case "ADA-W12WHT/BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = 6 * Panel_bool * removal_msbox
                    Case "ADA-W16BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-W16BRNxxx-xxx-x"
                        IO_Panel(kvp.Key) = 3 * Panel_bool * removal_msbox
                    Case "ADA-W18BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = Panel_bool * If(removal_msbox > 0, 3, 1)
                    Case "ADA-W18BRNxxx-xxx-x"
                        IO_Panel(kvp.Key) = Panel_bool * If(removal_msbox > 0, 3, 1)
                    Case "ADA-W18WHT/BLUxxx-xxx-x"
                        IO_Panel(kvp.Key) = Panel_bool * If(removal_msbox > 0, 3, 1)
                    Case "ADA-W16WHTxxx-xxx-x"
                        IO_Panel(kvp.Key) = 0
                    Case "ADA-W12GRNxxx-xxx-x"
                        IO_Panel(kvp.Key) = 3 * Panel_bool * removal_msbox

                    Case "ADA-W10EYE-YEL"
                        IO_Panel(kvp.Key) = 2 * Panel_bool * removal_msbox
                    Case "ADA-W10FERRULE-YEL"
                        IO_Panel(kvp.Key) = 0

                    Case "ADA-W12FERRULE-GRY"
                        IO_Panel(kvp.Key) = 4 * Panel_bool * removal_msbox
                    Case "ADA-W14FERRULE-BLU"
                        IO_Panel(kvp.Key) = 4 * Panel_bool * removal_msbox
                    Case "ADA-W16FERRULE-BLK"
                        IO_Panel(kvp.Key) = 20 * Panel_bool * removal_msbox
                    Case "ADA-W18FERRULE-RED"
                        IO_Panel(kvp.Key) = Panel_bool * (((IO_Panel.Item("ADA-SARB081-5m-LED") + IO_Panel.Item("ADA-SARB082-5m-LED")) * 3) + 30)
                    Case "ADA-W22FERRULE-WHT"
                        IO_Panel(kvp.Key) = Panel_bool * ((3 * ((IO_Panel.Item("ADA-SARB081-5m-LED") * 8) + (IO_Panel.Item("ADA-SARB082-5m-LED") * 16))) + 10)
                    Case "ADA-WIRE-WAY"
                        IO_Panel(kvp.Key) = Panel_bool * ((IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 3) * removal_msbox
                    Case "ADA-WIRE-WAY-Cover"
                        IO_Panel(kvp.Key) = Panel_bool * ((IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 3)
                    Case "ADA-DIN-RAIL"
                        IO_Panel(kvp.Key) = Panel_bool * If(removal_msbox > 0, 0, 1)
                    Case "ADA-DIN-RAIL-Tall"
                        IO_Panel(kvp.Key) = Panel_bool * ((IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) / 3)
                    Case "ADA-WIRE-TIE"
                        IO_Panel(kvp.Key) = Panel_bool * ((IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * 5)
                    Case "ADA-WIRE-TIE-MOUNT"
                        IO_Panel(kvp.Key) = Panel_bool * ((IO_Panel.Item("ADA-NEMA12-BOX24") + IO_Panel.Item("ADA-NEMA4X-BOX24")) * 5)


                        '------------------ Label parts ------------------
                    Case "ADA-WL (WireLabel)"
                        IO_Panel(kvp.Key) = 0 '0.2 * Panel_bool
                    Case "ADA-GND-L (Label)"
                        IO_Panel(kvp.Key) = 0 '0.06 * Panel_bool
                    Case "ADA-DECAL-BLK"
                        IO_Panel(kvp.Key) = Panel_bool

                        '-------------------Tools ---------------
                    Case "M12-M8TorqueDriver"
                        IO_Panel(kvp.Key) = 1 'If(ADA_Setup.ADA_set_list.count = 1, 1, 0)
                    Case "M12TorqueBit"
                        IO_Panel(kvp.Key) = 1
                    Case "M8TorqueBit"
                        IO_Panel(kvp.Key) = If(Panel_bool > 0, (If(Remote_IO < 1, Panel_bool, 0) / 40), 0)

                End Select
            Next
        Next
    End Sub

    'This will calculate the ADA parts in the PLC Panel (PLC Panel tab)


    Public Sub Calculate_PLC_Panel()

        Dim Logix_n As Double : Logix_n = ControlLogix_PLC_Box + CompactLogix_PLC_Box

        For i = 1 To 3
            For Each kvp As KeyValuePair(Of String, Double) In PLC_Panel.ToArray

                Select Case kvp.Key
                    Case "1756-A4"
                        PLC_Panel(kvp.Key) = Logix_n * (ControlLogix_PLC_Box * Logix_n)
                    Case "1756-A7"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-A10"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-A13"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-A17"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-PB72"
                        PLC_Panel(kvp.Key) = Logix_n * (ControlLogix_PLC_Box * Logix_n)
                    Case "1756-PB75"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-L71"
                        PLC_Panel(kvp.Key) = Logix_n * ControlLogix_PLC_Box
                    Case "1756-L72"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-DNB"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-EN2T"
                        PLC_Panel(kvp.Key) = Logix_n * (ControlLogix_PLC_Box * Logix_n)
                    Case "1756-HSC"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-CN2"
                        PLC_Panel(kvp.Key) = 0
                    Case "1756-N2"
                        PLC_Panel(kvp.Key) = (((PLC_Panel.Item("1756-A4") * 4) + (PLC_Panel.Item("1756-A7") * 7) + (PLC_Panel.Item("1756-A10") * 10) + (PLC_Panel.Item("1756-A13") * 13) + (PLC_Panel.Item("1756-A17") * 17)) -
                           (PLC_Panel.Item("1756-L71") + PLC_Panel.Item("1756-L72") + PLC_Panel.Item("1756-DNB") + PLC_Panel.Item("1756-EN2T") + PLC_Panel.Item("1756-HSC") + PLC_Panel.Item("1756-CN2"))) * Logix_n
                    Case "1756-BATM"
                        PLC_Panel(kvp.Key) = 0
                    Case "1769-ECR"
                        PLC_Panel(kvp.Key) = (PLC_Panel.Item("1756-L71") + PLC_Panel.Item("1756-L72")) * Logix_n
                    Case "1769-PB2"
                        PLC_Panel(kvp.Key) = Logix_n * PLC_Panel.Item("1769-L30ER")
                    Case "1769-PB4"
                        PLC_Panel(kvp.Key) = Logix_n * PLC_Panel.Item("1769-L33ER")
                    Case "1769-L16ER"
                        PLC_Panel(kvp.Key) = 0
                    Case "1769-L18ER"
                        PLC_Panel(kvp.Key) = 0
                    Case "1769-L24ER"
                        PLC_Panel(kvp.Key) = 0
                    Case "1769-L30ER"
                        PLC_Panel(kvp.Key) = 0
                    Case "1769-L33ER"
                        PLC_Panel(kvp.Key) = Logix_n * CompactLogix_PLC_Box
                    Case "1786-XT"
                        PLC_Panel(kvp.Key) = 0
                    Case "1786-TPR"
                        PLC_Panel(kvp.Key) = 0
                    Case "807966 ADATA"
                        PLC_Panel(kvp.Key) = Logix_n * PLC_Panel.Item("1769-L33ER")
                    Case "ADA-ETH-CELL-Router"
                        PLC_Panel(kvp.Key) = Logix_n * (PLC_Panel.Item("1786-XT") + PLC_Panel.Item("1786-TPR"))
                    Case "ADA-NEMA12-BOX24"
                        PLC_Panel(kvp.Key) = Logix_n * If(Logix_n And NEMA12 And (PLC_Panel.Item("1756-A7") + PLC_Panel.Item("1756-A10") + PLC_Panel.Item("1756-A13") + PLC_Panel.Item("1756-A17")) = 0, 1, 0)
                    Case "ADA-NEMA4X-BOX24"
                        PLC_Panel(kvp.Key) = Logix_n * If(Logix_n And NEMA4X And (PLC_Panel.Item("1756-A7") + PLC_Panel.Item("1756-A10") + PLC_Panel.Item("1756-A13") + PLC_Panel.Item("1756-A17")) = 0, 1, 0)
                    Case "ADA-SP12-21"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-PSL-PLC"
                        PLC_Panel(kvp.Key) = Logix_n * (PLC_Panel.Item("ADA-NEMA12-BOX24") + PLC_Panel.Item("ADA-NEMA4X-BOX24"))
                    Case "ADA-BOX-MOUNT-HW-KIT"
                        PLC_Panel(kvp.Key) = Logix_n * (PLC_Panel.Item("ADA-NEMA12-BOX24") + PLC_Panel.Item("ADA-NEMA4X-BOX24"))
                    Case "ADA-MOUNT-PAD"
                        PLC_Panel(kvp.Key) = 4 * Logix_n * Logix_n * (PLC_Panel.Item("ADA-NEMA12-BOX24") + PLC_Panel.Item("ADA-NEMA4X-BOX24"))
                    Case "ADA-EN-PB-NO-B"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-EN-PB-LIT"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-EN-PB-Legend"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-DIN-RAIL"
                        PLC_Panel(kvp.Key) = 0.25 * Logix_n
                    Case "ADA-6x10-W-TLBL"
                        PLC_Panel(kvp.Key) = Logix_n * PLC_Panel.Item("ADA-NEMA12-BOX24")
                    Case "ADA-1EB10"
                        PLC_Panel(kvp.Key) = Logix_n * Logix_n * PLC_Panel.Item("ADA-NEMA12-BOX24")
                    Case "ADA-1CB-02"
                        PLC_Panel(kvp.Key) = Logix_n * Logix_n * Logix_n * PLC_Panel.Item("ADA-NEMA12-BOX24")
                    Case "ADA-5x9-W-TLBL"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-1TB4"
                        PLC_Panel(kvp.Key) = 0 '8 * Logix_n
                    Case "ADA-1TBG"
                        PLC_Panel(kvp.Key) = 2 * Logix_n * (Logix_n * Logix_n * Logix_n * PLC_Panel.Item("ADA-NEMA12-BOX24"))
                    Case "ADA-1TEA"
                        PLC_Panel(kvp.Key) = 2 * Logix_n
                    Case "ADA-UPS-10-24"
                        PLC_Panel(kvp.Key) = PLC_UPS_Battery_Backup
                    Case "ADA-DECAL-BLK"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-WL (WireLabel)"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-RJ45-SWT-8-port"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-RJ45-SWT-16-port"
                        PLC_Panel(kvp.Key) = Logix_n
                    Case "ADA-ETH-CELL-Router"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-BXPCG-3/4"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-BXPCG-3/4R"
                        PLC_Panel(kvp.Key) = 6 * Logix_n
                    Case "ADA-BXPGN-3/4"
                        PLC_Panel(kvp.Key) = 6 * Logix_n
                    Case "ADA-BXORING-3/4"
                        PLC_Panel(kvp.Key) = 6 * Logix_n
                    Case "ADA-BXMHCG-0X0"
                        PLC_Panel(kvp.Key) = 6 * Logix_n
                    Case "ADA-BXMHCG-4X5"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-BXMHCG-4X6.5"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-BXMHCG-5X5"
                        PLC_Panel(kvp.Key) = 0
                    Case "ADA-BXMHCG-2X8"
                        PLC_Panel(kvp.Key) = 6 * Logix_n
                End Select
            Next
        Next
    End Sub

    'This will calculate the ADA parts in the Field (Field Parts tab)
    Public Sub Calculate_Field_Panel()

        Dim sum_f17_28 As Double : sum_f17_28 = 0

        For i = 0 To 11
            sum_f17_28 = sum_f17_28 + CType(Me.Motors_table(i, 0), Double)
        Next
        '------------------------
        Dim sum_f29_173 As Double : sum_f29_173 = 0

        For i = 12 To 47
            sum_f29_173 = sum_f29_173 + CType(Me.Motors_table(i, 0), Double)
        Next

        sum_f29_173 = sum_f29_173 + ADA_NEMA12_BOX24_No_Starters
        sum_f29_173 = sum_f29_173 + ADA_NEMA12_BOX24_No_Starters

        For i = 0 To 18
            sum_f29_173 = sum_f29_173 + CType(Me.inputs_table(i, 0), Double)
        Next

        For i = 0 To 18
            sum_f29_173 = sum_f29_173 + CType(Me.push_table(i, 0), Double)
        Next

        For i = 0 To 14
            sum_f29_173 = sum_f29_173 + CType(Me.lights_table(i, 0), Double)
        Next

        For i = 0 To 11
            sum_f29_173 = sum_f29_173 + CType(Me.field_table(i, 0), Double)
        Next

        For i = 0 To 11
            sum_f29_173 = sum_f29_173 + CType(Me.bus_table(i, 0), Double)
        Next

        For i = 0 To 3
            sum_f29_173 = sum_f29_173 + CType(Me.scanner_table(i, 0), Double)
        Next

        For i = 0 To 7
            sum_f29_173 = sum_f29_173 + CType(Me.brakes_table(i, 0), Double)
        Next

        '-------------------------
        Dim sum_f29_67 As Double : sum_f29_67 = 0

        For i = 12 To 47
            sum_f29_67 = sum_f29_67 + CType(Me.Motors_table(i, 0), Double)
        Next



        For i = 1 To 5
            For Each kvp As KeyValuePair(Of String, Double) In Field.ToArray

                Select Case kvp.Key
                    Case "ADA-ERP-Cable"
                        Field(kvp.Key) = Panel_bool * 1.05 * CType(Me.inputs_table(3, 0), Double)
                    Case "ADA-ERP-THIMBLE"
                        Field(kvp.Key) = If((Panel_bool * 1.05 * CType(Me.inputs_table(3, 0), Double)) > 0, (Panel_bool * CType(Me.inputs_table(1, 0), Double) * 2) + (Panel_bool * CType(Me.inputs_table(2, 0), Double) * 4), 0) * Panel_bool
                    Case "ADA-ERP-EYEBOLY"
                        Field(kvp.Key) = Math.Ceiling((Panel_bool * 1.05 * CType(Me.inputs_table(3, 0), Double)) / 100) * Panel_bool
                    Case "ADA-ERP-TB"
                        Field(kvp.Key) = If((Panel_bool * 1.05 * CType(Me.inputs_table(3, 0), Double)) > 0, (Panel_bool * CType(Me.inputs_table(1, 0), Double)) + (Panel_bool * CType(Me.inputs_table(2, 0), Double) * 2), 0) * Panel_bool
                    Case "ADA-ERP-1-VDC-H"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(1, 0), Double)
                    Case "ADA-ERP-2-VDC-H"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(2, 0), Double)
                    Case "ADA-ERP-1-VDC-S"
                        Field(kvp.Key) = 0
                    Case "ADA-ERP-Pulley-Kit"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(4, 0), Double)
                    Case "ADA-BXPCG-1/2"
                        Field(kvp.Key) = Panel_bool * (Field.Item("ADA-ERP-1-VDC-H") + Field.Item("ADA-ERP-2-VDC-H"))
                    Case "ADA-BXPCP-PG7"
                        Field(kvp.Key) = 1
                    Case "ADA-SAC4-00-M"
                        Field(kvp.Key) = 1
                    Case "ADA-FTERM"
                        Field(kvp.Key) = Panel_bool * (CType(Me.inputs_table(1, 0), Double) + CType(Me.inputs_table(2, 0), Double) + CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(8, 0), Double))
                    Case "ADA-E-STOP PATENT LABEL"
                        Field(kvp.Key) = Panel_bool * ((CType(Me.inputs_table(1, 0), Double) * 2) + CType(Me.inputs_table(2, 0), Double) + (CType(Me.inputs_table(5, 0), Double) * 2) + (CType(Me.inputs_table(6, 0), Double) * 2))

                        '-------------------- INTERLOCK BOXES --------------
                    Case "ADA-100/75BX-W"
                        Field(kvp.Key) = 0  'Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-100/100BX-G"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-100/100BX-TP"
                        Field(kvp.Key) = NEMA4X * CType(Me.inputs_table(9, 0), Double)
                    Case "ADA-150/150BX-G"
                        Field(kvp.Key) = PLC_GateWay_PGW * CType(Me.inputs_table(10, 0), Double)
                    Case "ADA-150/150BX-TP"
                        Field(kvp.Key) = ControlLogix_PLC_Box * CType(Me.inputs_table(11, 0), Double)
                    Case "ADA-IDEC-BASE"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * (CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double))
                    Case "ADA-IDEC-24R"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * (CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double))
                    Case "ADA-1TEA"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-1TB2_2"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * CType(Me.inputs_table(7, 1), Double)
                    Case "ADA-5x9-W-TLBL"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * CType(Me.inputs_table(7, 1), Double) * 2
                    Case "ADA-1EB2"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-1TB-JP2"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * CType(Me.inputs_table(7, 1), Double) / 2
                    Case "ADA-DIN-RAIL-INTERL"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) / 20
                    Case "ADA-SAC5-03-G"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double) * (CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2
                    Case "ADA-BXPCG-3/4R"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-BXPGN-3/4"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(7, 0), Double)
                    Case "ADA-BXMHCG-4X5"
                        Field(kvp.Key) = If(((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) = 3, 1, 0) * Field.Item("ADA-100/75BX-W")
                    Case "ADA-BXMHCG-4X6.5"
                        Field(kvp.Key) = If(((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) = 4, 1, 0) * Field.Item("ADA-100/75BX-W")
                    Case "ADA-BXMHCG-5X5"
                        Field(kvp.Key) = If(((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) = 5, 1, 0) * Field.Item("ADA-100/75BX-W")
                    Case "ADA-BXMHCG-2X8"
                        Field(kvp.Key) = If(((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) = 2, 1, 0) * Field.Item("ADA-100/75BX-W")
                    Case "ADA-PTF4-05-G"
                        Field(kvp.Key) = If(((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) * CType(Me.inputs_table(7, 0), Double) > 0, ((CType(Me.inputs_table(7, 1), Double) + CType(Me.inputs_table(7, 2), Double) + CType(Me.inputs_table(7, 3), Double)) / 2) * CType(Me.inputs_table(7, 0), Double), 0)

                        '------------------ Encoder -------------
                    Case "ADA-QUAD-ENCODER"
                        Field(kvp.Key) = 0 'R2_setup * CType(Me.field_table(9, 0), Double)
                    Case "ADA-ENCODER"
                        Field(kvp.Key) = Panel_bool * CType(Me.field_table(10, 0), Double)
                    Case "ADA-SAC8-10-B-SH"
                        Field(kvp.Key) = Field.Item("ADA-QUAD-ENCODER") + Field.Item("ADA-ENCODER")
                    Case "ADA-SAC8-20-B-SH"
                        Field(kvp.Key) = 0
                    Case "ADA-SAC8-M-SH"
                        Field(kvp.Key) = Field.Item("ADA-QUAD-ENCODER") + Field.Item("ADA-ENCODER")

                        '------------------- Air pressure ---------------
                    Case "ADA-PS"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(8, 0), Double)
                    Case "ADA-SAC4-03-M-m12-F-m8"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(8, 0), Double)
                    Case "ADA-BHR5-M-M20"
                        Field(kvp.Key) = NEMA12 * CType(Me.inputs_table(9, 0), Double)

                        '-------------------- Prox switch ---------------
                    Case "ADA-PX"
                        Field(kvp.Key) = CType(Me.inputs_table(9, 0), Double) * Panel_bool

                        '--------------------- Limit switches and accesories -----------
                    Case "ADA-LS"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(10, 0), Double)
                    Case "ADA-BHR5-M-M20"
                        Field(kvp.Key) = NEMA12 * CType(Me.inputs_table(12, 0), Double)
                    Case "ADA-LS-ARM"
                        Field(kvp.Key) = Panel_bool * Field.Item("ADA-LS")

                        '-------------------------- Photoeyes and accesories -----------
                    Case "ADA-PE-RR-2HOLE"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(11, 0), Double)
                    Case "ADA-PE-RR-2HOLE-MNT"
                        Field(kvp.Key) = Panel_bool * Panel_bool * CType(Me.inputs_table(11, 0), Double) + (Panel_bool * CType(Me.inputs_table(11, 0), Double))
                    Case "ADA-PE-REF-2HOLE"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(11, 0), Double)
                    Case "ADA-PE-RR-18mm"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(12, 0), Double)
                    Case "ADA-PE-ER-18mm"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(13, 0), Double)
                    Case "ADA-PE-RE-18mm"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(13, 0), Double)
                    Case "ADA-PE-DF-150"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(14, 0), Double)
                    Case "ADA-PE-DF-450"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(15, 0), Double)
                    Case "ADA-PE-DF-1000"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(17, 0), Double)
                    Case "ADA-PE-DF-1000MNT"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(17, 0), Double)
                    Case "ADA-PX-US"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(17, 0), Double)
                    Case "ADA-HY-EZ-AUX"
                        Field(kvp.Key) = CType(Me.field_table(1, 0), Double) + CType(Me.field_table(2, 0), Double)
                    Case "ADA-SAC4-00-F-F"
                        Field(kvp.Key) = If(CType(Me.push_table(18, 0), Double) > 1, 1, 0) + If(CType(Me.push_table(18, 0), Double) > 50, 1, 0) + If(CType(Me.push_table(18, 0), Double) > 100, 1, 0) + If(CType(Me.push_table(18, 0), Double) > 150, 1, 0)
                    Case "ADA-PE-REF"
                        Field(kvp.Key) = Field.Item("ADA-PE-RR-18mm")
                    Case "ADA-PE-MNT-KIT"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(18, 0), Double)
                    Case "ADA-PE-HW-KIT"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(18, 0), Double)

                        '----------------- Pushbuttons ----------------
                    Case "ADA-1PBX-WHT"
                        Field(kvp.Key) = CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "ADA-1PBX-YEL"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double)
                    Case "ADA-1PBX-YEL-GUARD"
                        Field(kvp.Key) = CType(Me.inputs_table(6, 0), Double)
                    Case "ADA-2PBX-WHT"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double)
                    Case "ADA-3PBX-WHT"
                        Field(kvp.Key) = CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "ADA-PB-RD-MUSH"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double)
                    Case "ADA-PB-RED"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double)
                    Case "ADA-PB-GRN"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double)
                    Case "ADA-PB-BLU"
                        Field(kvp.Key) = CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double) + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(8, 0), Double)
                    Case "ADA-PB-BLU-LT"
                        Field(kvp.Key) = CType(Me.push_table(7, 0), Double)
                    Case "ADA-PB-BLK"
                        Field(kvp.Key) = CType(Me.push_table(2, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "ADA-PB-CLR-LT"
                        Field(kvp.Key) = 0
                    Case "ADA-SWT2-BLK"
                        Field(kvp.Key) = CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "ADA-SWT2-KEY"
                        Field(kvp.Key) = CType(Me.push_table(5, 0), Double)
                    Case "ADA-PB-MNT"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + (CType(Me.push_table(0, 0), Double) * 2) + (CType(Me.push_table(1, 0), Double) * 3) +
                            (CType(Me.push_table(2, 0), Double) * 3) + (CType(Me.push_table(3, 0), Double) * 3) + (CType(Me.push_table(4, 0), Double) * 3) + (CType(Me.push_table(5, 0), Double) * 3) _
                            + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "ADA-PB-NO"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + (CType(Me.push_table(0, 0), Double) * 1) + (CType(Me.push_table(1, 0), Double) * 2) +
                            (CType(Me.push_table(2, 0), Double) * 2) + (CType(Me.push_table(3, 0), Double) * 2) + (CType(Me.push_table(4, 0), Double) * 3) + (CType(Me.push_table(5, 0), Double) * 4) _
                            + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "ADA-PB-NC"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + (CType(Me.push_table(0, 0), Double) * 1) + (CType(Me.push_table(1, 0), Double) * 1) +
                            (CType(Me.push_table(2, 0), Double) * 1) + (CType(Me.push_table(3, 0), Double) * 2) + (CType(Me.push_table(4, 0), Double) * 1) + (CType(Me.push_table(5, 0), Double) * 1)
                    Case "ADA-PB-Stop"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double)
                    Case "ADA-PB-Start"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double)
                    Case "ADA-PB-Reset"
                        Field(kvp.Key) = CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double)
                    Case "ADA-PB-Auto-Man"
                        Field(kvp.Key) = CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "ADA-PB-Jog"
                        Field(kvp.Key) = CType(Me.push_table(4, 0), Double)
                    Case "ADA-PB-Home"
                        Field(kvp.Key) = CType(Me.push_table(5, 0), Double)
                    Case "ADA-PB-Left-Right"
                        Field(kvp.Key) = CType(Me.push_table(5, 0), Double)
                    Case "ADA-22x22-S-PBLBL"
                        Field(kvp.Key) = (CType(Me.push_table(0, 0), Double) * 2) + (CType(Me.push_table(1, 0), Double) * 3) + (CType(Me.push_table(3, 0), Double) * 3) + (CType(Me.push_table(4, 0), Double) * 3) +
                            (CType(Me.push_table(5, 0), Double) * 3) + (CType(Me.push_table(6, 0), Double)) + (CType(Me.push_table(7, 0), Double)) + (CType(Me.push_table(8, 0), Double)) + (CType(Me.push_table(9, 0), Double))
                    Case "ADA-22x22-W-PBLBL"
                        Field(kvp.Key) = (CType(Me.push_table(7, 0), Double))
                    Case "ADA-1LT-LED-RED"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double)
                    Case "ADA-1LT-LED-YEL"
                        Field(kvp.Key) = 0
                    Case "ADA-1LT-LED-GRN"
                        Field(kvp.Key) = 0
                    Case "ADA-1LT-LED-WHT"
                        Field(kvp.Key) = 0
                    Case "ADA-1LT-LED-BLU"
                        Field(kvp.Key) = (CType(Me.push_table(7, 0), Double))
                    Case "ADA-SARC012-m12-0.3M-B_FMM"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "ADA-SAC5-05-Y"
                        Field(kvp.Key) = CType(Me.inputs_table(1, 0), Double) + CType(Me.inputs_table(2, 0), Double) + CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(2, 0), Double) + (CType(Me.push_table(8, 0), Double))
                    Case "ADA-PBC4-03-G"
                        Field(kvp.Key) = CType(Me.inputs_table(1, 0), Double) + CType(Me.inputs_table(2, 0), Double) + CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(0, 0), Double) + (CType(Me.push_table(1, 0), Double) * 2) +
                             (CType(Me.push_table(2, 0), Double)) + (CType(Me.push_table(3, 0), Double) * 2) + (CType(Me.push_table(4, 0), Double) * 2) + (CType(Me.push_table(5, 0), Double) * 2) + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "ADA-BXPGN-M20"
                        Field(kvp.Key) = CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double) +
                            CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double) +
                            CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double)
                    Case "ADA-BXPCG-M20R"
                        Field(kvp.Key) = CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double) +
                            CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "ADA-BXMHCGM20-2X5"
                        Field(kvp.Key) = CType(Me.push_table(1, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double) + CType(Me.push_table(8, 0), Double)
                    Case "ADA-BXMHCGM20-3X5"
                        Field(kvp.Key) = CType(Me.inputs_table(1, 0), Double) + CType(Me.inputs_table(2, 0), Double) + CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double) + CType(Me.push_table(2, 0), Double)
                    Case "ADA-22x22-S-PBLBL"
                        Field(kvp.Key) = (2 * (CType(Me.push_table(0, 0), Double) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double)) + CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) +
                            CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double) + CType(Me.push_table(5, 0), Double)) / 88
                    Case "ADA-W22BLUxxx-xxx-x"
                        Field(kvp.Key) = ((Field.Item("ADA-ERP-1-VDC-H") + Field.Item("ADA-ERP-2-VDC-H") + Field.Item("ADA-ERP-1-VDC-S")) * 13) + ((Field.Item("ADA-1PBX-YEL") + Field.Item("ADA-1PBX-YEL-GUARD")) * 13) + (Field.Item("ADA-100/75BX-W") * 20) + (Field.Item("ADA-QUAD-ENCODER") * 8) + (Field.Item("ADA-LS") * 4) + (Field.Item("ADA-1PBX-WHT") * 5) +
                            ((Field.Item("ADA-2PBX-WHT") + Field.Item("ADA-3PBX-WHT")) * 15) + (Field.Item("ADA-PB-Jog") * 9) + (Field.Item("ADA-STK-LT-BASE") * 6)
                    Case "ADA-W22FERRULE-WHT"
                        Field(kvp.Key) = ((Field.Item("ADA-ERP-1-VDC-H") + Field.Item("ADA-ERP-2-VDC-H") + Field.Item("ADA-ERP-1-VDC-S")) * 13) + ((Field.Item("ADA-1PBX-YEL") + Field.Item("ADA-1PBX-YEL-GUARD")) * 13) + (Field.Item("ADA-100/75BX-W") * 20) + (Field.Item("ADA-QUAD-ENCODER") * 8) + (Field.Item("ADA-LS") * 4) + (Field.Item("ADA-1PBX-WHT") * 5) +
                            ((Field.Item("ADA-2PBX-WHT") + Field.Item("ADA-3PBX-WHT")) * 15) + (Field.Item("ADA-PB-Jog") * 9) + (Field.Item("ADA-STK-LT-BASE") * 6)
                    Case "ADA-WL (WireLabel)"
                        Field(kvp.Key) = Math.Ceiling((((Field.Item("ADA-ERP-1-VDC-H") + Field.Item("ADA-ERP-2-VDC-H") + Field.Item("ADA-ERP-1-VDC-S")) * 4) + ((Field.Item("ADA-1PBX-YEL") + Field.Item("ADA-1PBX-YEL-GUARD")) * 4) + (Field.Item("ADA-100/75BX-W") * 6) +
                        ((Field.Item("ADA-QUAD-ENCODER") + Field.Item("ADA-ENCODER")) * 2) + (Field.Item("ADA-LS") * 2) + ((Field.Item("ADA-LS-ARM") + Field.Item("ADA-PE-RR-2HOLE") + Field.Item("ADA-PE-RR-2HOLE-MNT") + Field.Item("ADA-PE-REF-2HOLE") + Field.Item("ADA-PE-RR-18mm") + Field.Item("ADA-PE-ER-18mm") + Field.Item("ADA-PE-RE-18mm") + Field.Item("ADA-PE-DF-150") +
                        Field.Item("ADA-PE-DF-450") + Field.Item("ADA-PE-DF-1000") + Field.Item("ADA-PE-DF-1000MNT") + Field.Item("ADA-HY-EZ-AUX") + Field.Item("ADA-SAC4-00-F-F") + Field.Item("ADA-PE-REF")) * 2) + (Field.Item("ADA-1PBX-WHT") * 2) +
                        (Field.Item("ADA-STK-LT-BASE") * 2) + (Field.Item("ADA-STK-LT-BASE") * 2) + ((Field.Item("ADA-SARC021-m12-1M-MFF") + Field.Item("ADA-SARC012-m12-1M-FMM") + Field.Item("ADA-SAC5-00-M")) * 2) + (Field.Item("ADA-PBC4-03-G") * 3)) / 200)

                        '----------- Motor disconnect for > 50 ft ---------
                    Case "ADA-MD-KIT"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double)
                        '---------------- 50mm Stack-Lights and Accessories (70mm optional from P&F) ------
                    Case "ADA-STK-LT-RED"
                        Field(kvp.Key) = CType(Me.lights_table(8, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-YEL"
                        Field(kvp.Key) = CType(Me.lights_table(5, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-GRN"
                        Field(kvp.Key) = CType(Me.lights_table(7, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-BLU"
                        Field(kvp.Key) = CType(Me.lights_table(6, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-HRN-85"
                        Field(kvp.Key) = CType(Me.lights_table(9, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-HRN-100-8TN"
                        Field(kvp.Key) = CType(Me.lights_table(10, 0), Double) * Panel_bool
                    Case "ADA-STK-LT-BASE"
                        Field(kvp.Key) = ADA_Setup.Max_number(Field.Item("ADA-STK-LT-RED"), Field.Item("ADA-STK-LT-YEL"), Field.Item("ADA-STK-LT-GRN"), Field.Item("ADA-STK-LT-BLU"), Field.Item("ADA-STK-LT-HRN-85"))
                    Case "ADA-STK-LT-POLE"
                        Field(kvp.Key) = ADA_Setup.Max_number(Field.Item("ADA-STK-LT-RED"), Field.Item("ADA-STK-LT-YEL"), Field.Item("ADA-STK-LT-GRN"), Field.Item("ADA-STK-LT-BLU"), Field.Item("ADA-STK-LT-HRN-85"))
                    Case "ADA-STK-LT-LED"
                        Field(kvp.Key) = 0
                    Case "ADA-SAC3-0.06-B-FMM-5x4"
                        Field(kvp.Key) = Math.Round((Field.Item("ADA-STK-LT-RED") + Field.Item("ADA-STK-LT-YEL") + Field.Item("ADA-STK-LT-GRN") + Field.Item("ADA-STK-LT-BLU") + Field.Item("ADA-STK-LT-HRN-85") + Field.Item("ADA-STK-LT-HRN-100-8TN")) / 3) * Panel_bool

                        '------------------- Solenoid accesories ---------
                    Case "ADA-PBC4-05-G"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-00-10mm"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-00-11mm"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-00-18mm"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-0.3-Y-9.4mmMac"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-0.3-G-9.4mm"
                        Field(kvp.Key) = 0
                    Case "ADA-SOLC-0.6-G-9.4mm"
                        Field(kvp.Key) = 0
                    Case "ADA-SARC021-m12-1M-MFF"
                        Field(kvp.Key) = Field.Item("ADA-IOM-X2X-8Port-16IO-x67CR") + Field.Item("ADA-IOM-SWD-8Port-16In-x67CR") + Field.Item("ADA-IOM-SWD-8Port-16IO-x67CR") + Field.Item("ADA-IOM-SWD-8Port-8In-8Motx67CR") + Field.Item("ADA-IOM-SWD-8Port-16Mot-x67CR") + Field.Item("ADA-SWD5-1PORT") + Field.Item("ADA-SWD5-2PORT") + Field.Item("ADA-IOM-EIP-8Port-IOLink-LED")
                    Case "ADA-SARC021-m12-2M-MFF"
                        Field(kvp.Key) = 0
                    Case "ADA-SARC021-m12-0-MFF"
                        Field(kvp.Key) = 0
                    Case "ADA-SARC012-m12-0-MFF"
                        Field(kvp.Key) = 0
                    Case "ADA-SARC012-m12-1M-FMM"
                        Field(kvp.Key) = CType(Me.field_table(2, 0), Double) + CType(Me.field_table(4, 0), Double) + If((CType(Me.lights_table(5, 0), Double) + CType(Me.lights_table(6, 0), Double) + CType(Me.lights_table(7, 0), Double) + CType(Me.lights_table(8, 0), Double) + CType(Me.lights_table(9, 0), Double) + CType(Me.lights_table(10, 0), Double)) > 1, 1, 0)
                    Case "ADA-SAC5-00-M"
                        Field(kvp.Key) = 0 'Math.Ceiling(If(R2_setup > 0, CType(Me.field_table(1, 0), Double) + CType(Me.field_table(2, 0), Double) + CType(Me.field_table(10, 0), Double), 0))
                    Case "ADA-SAC5-00-F"
                        Field(kvp.Key) = 0
                    Case "ADA-PBC5-05-G"
                        Field(kvp.Key) = 0
                    Case "ADA-BHR5-F-NPT-12"
                        Field(kvp.Key) = 0
                    Case "ADA-BHR5-M-NPT-12"
                        Field(kvp.Key) = 0
                    Case "ADA-BHR5-M-PG-9"
                        Field(kvp.Key) = 0
                    Case "ADA-BHR5-F-PG-9"
                        Field(kvp.Key) = 0

                        '----------------Misc inputs and accesories -------------------

                    Case "ADA-PBC4-03-G"
                        Field(kvp.Key) = If(Valve_solenoid_adapter = 0, CType(Me.lights_table(13, 0), Double), 0)
                    Case "ADA-SWD5-1PORT"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(2, 0), Double), 0) * removal_msbox
                    Case "ADA-SWD5-2PORT"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(3, 0), Double), 0) * removal_msbox
                    Case "ADA-SWD5-SWD-RIB-F"
                        Field(kvp.Key) = If((If(Panel_bool > 0, CType(Me.bus_table(2, 0), Double), 0) + If(Panel_bool > 0, CType(Me.bus_table(3, 0), Double), 0)) > 0, 1, 0) * removal_msbox
                    Case "ADA-SWD5-TERM"
                        Field(kvp.Key) = If((If(Panel_bool > 0, CType(Me.bus_table(2, 0), Double), 0) + If(Panel_bool > 0, CType(Me.bus_table(3, 0), Double), 0)) > 0, 1, 0) * removal_msbox
                    Case "ADA-IOM-X2X-x67COM-CBL-0.25m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67COM-CBL-10m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67COM-CBL-15m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67PWR-CBL-0.25m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67PWR-CBL-10m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67PWR-CBL-15m"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0), (If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0) + If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)), 0)
                    Case "ADA-IOM-X2X-x67PWR-Supply"
                        Field(kvp.Key) = If(If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0) > 0, 1, 0)
                    Case "ADA-IOM-X2X-8Port-16IO-x67CR"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(1, 0), Double), 0)
                    Case "ADA-IOM-SWD-8Port-16In-x67CR"
                        Field(kvp.Key) = 0
                    Case "ADA-IOM-SWD-8Port-16IO-x67CR"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double), 0) * removal_msbox
                    Case "ADA-IOM-SWD-8Port-8In-8Motx67CR"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(6, 0), Double), 0) * removal_msbox
                    Case "ADA-IOM-SWD-8Port-16Mot-x67CR"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(7, 0), Double), 0) * removal_msbox
                    Case "ADA-M12-MINI-1-G-PWR"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(7, 0), Double), 0)
                    Case "ADA-IOM-SWD-PF1-2-x67CR"
                        Field(kvp.Key) = (If(Panel_bool > 0, (((CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8) + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 4), 0) + (CType(Me.field_table(9, 0), Double) + CType(Me.field_table(10, 0), Double)) / 24) * removal_msbox
                    Case "ADA-IOM-SWD-ENCODER-x67CR"
                        Field(kvp.Key) = If(Remote_IO > 0, If(Panel_bool > 0, CType(Me.field_table(9, 0), Double) + CType(Me.field_table(10, 0), Double), 0), 0) * removal_msbox
                    Case "ADA-SAC4-05-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SAC4-10-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SAC4-20-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SAC4-30-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SAC4-40-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SAC4-50-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, (CType(Me.bus_table(3, 0), Double) + CType(Me.bus_table(4, 0), Double)) / 8 + (CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(8, 0), Double)) / 4, 0) + CType(Me.bus_table(7, 0), Double)) / 4
                    Case "ADA-SWD8-50-GRN"
                        Field(kvp.Key) = If(Panel_bool > 0, Math.Round(If(((sum_f17_28 / 10) + (sum_f29_173 / 5) + (MDRs / 50)) < 5, ((sum_f17_28 / 10) + (sum_f29_173 / 5) + (MDRs / 50)), 0), MidpointRounding.AwayFromZero), 0) * removal_msbox
                    Case "ADA-SWD8-250-GRN-MW"
                        Field(kvp.Key) = Math.Round(removal_msbox * (If(Panel_bool > 0, If(((sum_f17_28 / 10) + (sum_f29_173 / 5) + (MDRs / 50)) >= 5, ((sum_f17_28 / 50) + (sum_f29_67 / 25) + (MDRs / 250)), 0), 0))) * removal_msbox
                    Case "ADA-IOM-EIP-8Port-IOLink-LED"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + ((Field.Item("ADA-IOM-IOLink-8Port-16IN-SLAVE") + Field.Item("ADA-IOM-IOLink-8Port-10IN6OUT-SLAVE")) / 7), 0)
                    Case "ADA-IOM-IOLink-8Port-16IN-SLAVE"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(10, 0), Double), 0)
                    Case "ADA-IOM-IOLink-8Port-10IN6OUT-SLAVE"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(11, 0), Double), 0)
                    Case "ADA-IOM-IOLink-8Port-16IN16OUT-SLAVE"
                        Field(kvp.Key) = 0
                    Case "ADA-SAC5-0,6-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + Starter_Panel.Item("ADA-SAC5-0,5-G-PWR-M-Bulkhead") / 4, 0))
                    Case "ADA-SAC5-03-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + Starter_Panel.Item("ADA-SAC5-0,5-G-PWR-M-Bulkhead") / 4, 0))
                    Case "ADA-SAC5-05-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + Starter_Panel.Item("ADA-SAC5-0,5-G-PWR-M-Bulkhead") / 4, 0)) + Panel_bool
                    Case "ADA-SAC5-10-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + Starter_Panel.Item("ADA-SAC5-0,5-G-PWR-M-Bulkhead") / 4, 0))
                    Case "ADA-SAC5-15-G-PWR"
                        Field(kvp.Key) = Math.Ceiling(If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double) + Starter_Panel.Item("ADA-SAC5-0,5-G-PWR-M-Bulkhead") / 4, 0))
                    Case "ADA-EIP4-2-GRN"
                        Field(kvp.Key) = 0
                    Case "ADA-EIP4-4-GRN"
                        Field(kvp.Key) = 0
                    Case "ADA-EIP4-6-GRN"
                        Field(kvp.Key) = If(Panel_bool > 0, CType(Me.bus_table(9, 0), Double), 0)
                    Case "ADA-EIP4-10-GRN"
                        Field(kvp.Key) = 0
                    Case "ADA-EIP4-15-GRN"
                        Field(kvp.Key) = 0
                    Case "ADA-EIP4-20-GRN"
                        Field(kvp.Key) = 0

                        '----------------- Mouting hardware option --------
                    Case "ADA-UNI-MNT-FLAT-SM"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(4, 0), Double)
                    Case "ADA-UNI-MNT-45-SM"
                        Field(kvp.Key) = If(include_ADA_brackets > 0, 2 * Field.Item("ADA-1PBX-WHT") + Field.Item("ADA-1PBX-YEL") + Field.Item("ADA-1PBX-YEL-GUARD") + Field.Item("ADA-2PBX-WHT"), 0)
                    Case "ADA-UNI-MNT-90-SM"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(4, 0), Double)
                    Case "ADA-UNI-MNT-FLAT-LG"
                        Field(kvp.Key) = If(include_ADA_brackets > 0, Panel_bool * CType(Me.inputs_table(2, 0), Double), 0) * NEMA12
                    Case "ADA-UNI-MNT-45-LG"
                        Field(kvp.Key) = If(include_ADA_brackets > 0, (Field.Item("ADA-PB-Start") + Field.Item("ADA-PB-Reset") + Field.Item("ADA-PB-Auto-Man") + Field.Item("ADA-PB-Jog") + Field.Item("ADA-3PBX-WHT") + Field.Item("ADA-100/75BX-W")) * Panel_bool + (IO_Panel.Item("ADA-SARB081-5m-LED") + IO_Panel.Item("ADA-SARB082-5m-LED") + IO_Panel.Item("ADA-SARB081-m23P12") + IO_Panel.Item("ADA-SARB042-m23P19-LED") + IO_Panel.Item("ADA-SARB082-m23P19-LED") + IO_Panel.Item("ADA-SARB161-m23P19-LED")), 0) +
                            CType(Me.bus_table(0, 0), Double) + CType(Me.bus_table(1, 0), Double) + CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(6, 0), Double) + CType(Me.bus_table(7, 0), Double)
                    Case "ADA-UNI-MNT-90-LG"
                        Field(kvp.Key) = If(include_ADA_brackets > 0, Panel_bool * CType(Me.inputs_table(2, 0), Double), 0) * Panel_bool
                    Case "ADA-UNI-MNT-HW-KIT"
                        Field(kvp.Key) = Field.Item("ADA-UNI-MNT-FLAT-SM") + Field.Item("ADA-UNI-MNT-45-SM") + Field.Item("ADA-UNI-MNT-90-SM") + Field.Item("ADA-UNI-MNT-FLAT-LG") + Field.Item("ADA-UNI-MNT-45-LG") + Field.Item("ADA-UNI-MNT-90-LG")
                    Case "M12-M8TorqueDriver"
                        Field(kvp.Key) = If(Panel_bool > 0, ((CType(Me.bus_table(2, 0), Double) + CType(Me.bus_table(3, 0), Double)) / 80) + ((CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 40) + ((CType(Me.bus_table(7, 0), Double)) / 40), 0)
                    Case "M12TorqueBit"
                        Field(kvp.Key) = If(Panel_bool > 0, ((CType(Me.bus_table(2, 0), Double) + CType(Me.bus_table(3, 0), Double)) / 80) + ((CType(Me.bus_table(4, 0), Double) + CType(Me.bus_table(5, 0), Double) + CType(Me.bus_table(7, 0), Double)) / 40) + ((CType(Me.bus_table(7, 0), Double)) / 40), 0)
                    Case "M8TorqueBit"
                        Field(kvp.Key) = 0

                        '---------- labor -----------
                    Case "2PBX Labor"
                        Field(kvp.Key) = CType(Me.push_table(0, 0), Double)
                    Case "3PBX Labor"
                        Field(kvp.Key) = CType(Me.push_table(1, 0), Double) + CType(Me.push_table(2, 0), Double) + CType(Me.push_table(3, 0), Double) + CType(Me.push_table(4, 0), Double) + CType(Me.push_table(5, 0), Double)
                    Case "1PBX E-Stop Labor"
                        Field(kvp.Key) = CType(Me.inputs_table(5, 0), Double) + CType(Me.inputs_table(6, 0), Double)
                    Case "1PBX Labor"
                        Field(kvp.Key) = CType(Me.push_table(6, 0), Double) + CType(Me.push_table(7, 0), Double) + CType(Me.push_table(8, 0), Double) + CType(Me.push_table(9, 0), Double)
                    Case "1ERP Labor"
                        Field(kvp.Key) = Panel_bool * CType(Me.inputs_table(1, 0), Double) + Panel_bool * CType(Me.inputs_table(2, 0), Double)

                End Select
            Next
        Next
    End Sub

    'This will calculate the ADA parts in the Scanner(Scanner Parts tab)
    Public Sub Calculate_Scanners()

        For Each kvp As KeyValuePair(Of String, Double) In Scanners.ToArray

                Select Case kvp.Key
                    Case "SICK Bar Code Scan 430 KIT"
                    Scanners(kvp.Key) = Panel_bool * CType(Me.scanner_table(0, 0), Double)
                Case "SICK Bar Code Scan 490 OMNI KIT"
                    Scanners(kvp.Key) = Panel_bool * CType(Me.scanner_table(1, 0), Double)
                Case "Cognex Bar Code Scan 430 KIT"
                    Scanners(kvp.Key) = Panel_bool * CType(Me.scanner_table(2, 0), Double)
                Case "Cognex Insite Vision KIT"
                    Scanners(kvp.Key) = Panel_bool * CType(Me.scanner_table(3, 0), Double)
            End Select
            Next

    End Sub

    'This will calculate M12 Cables qty
    'Public Sub Calculate_M12()


    '    For Each kvp As KeyValuePair(Of String, Double) In M12.ToArray

    '            Select Case kvp.Key
    '                Case "ADA-SAC5-02-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-SAC5-03-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-SAC5-05-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-SAC5-10-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-SAC5-15-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-SAC5-20-G"
    '                    M12(kvp.Key) = 0
    '                Case "ADA-23x4-W-WLBL"
    '                    M12(kvp.Key) = 0
    '            End Select
    '        Next

    'End Sub

    ''This will calculate M12 _ES Cables qty
    'Public Sub Calculate_M12_ES()


    '    For Each kvp As KeyValuePair(Of String, Double) In M12_ES.ToArray

    '            Select Case kvp.Key
    '                Case "ADA-SAC5-02-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-SAC5-03-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-SAC5-05-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-SAC5-10-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-SAC5-15-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-SAC5-20-Y"
    '                    M12_ES(kvp.Key) = 0
    '                Case "ADA-23x4-W-WLBL"
    '                    M12_ES(kvp.Key) = 0
    '            End Select
    '        Next

    'End Sub



End Class
