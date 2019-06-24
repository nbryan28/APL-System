Public Class Setup_total
    Private Sub Setup_total_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        setup_grid.Rows.Add(New String() {""})
        setup_grid.Rows.Add(New String() {"NEMA12"})
        setup_grid.Rows.Add(New String() {"NEMA4X"})
        setup_grid.Rows.Add(New String() {"PLC in GateWay, PGW"})
        setup_grid.Rows.Add(New String() {"ControlLogix PLC Box"})
        setup_grid.Rows.Add(New String() {"CompactLogix PLC Box"})
        setup_grid.Rows.Add(New String() {"PLC UPS (Battery Backup)"})
        setup_grid.Rows.Add(New String() {"Amp Mon. Electro Motor Starters"})
        setup_grid.Rows.Add(New String() {"Remote I/O Only"})
        setup_grid.Rows.Add(New String() {"Total Current Per Group"})
        setup_grid.Rows.Add(New String() {""})
        setup_grid.Rows.Add(New String() {"Motor Driven Rollers (MDRs)"})

        setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.25"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.50"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  0.75"})
        setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  1.00"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  1.50"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  2.00"})
        setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  2.50"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  3.00"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  4.00"})
        setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  5.00"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  7.50"}) : setup_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)  10.0"})

        setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.25"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.50"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  0.75"})
        setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  1.00"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  1.50"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  2.00"})
        setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  2.50"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  3.00"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  4.00"})
        setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  5.00"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  7.50"}) : setup_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)  10.0"})

        setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.25"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.50"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  0.75"})
        setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  1.00"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  1.50"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  2.00"})
        setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  2.50"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  3.00"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  4.00"})
        setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  5.00"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  7.50"}) : setup_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)  10.0"})

        setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.25"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.50"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  0.75"})
        setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  1.00"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  1.50"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  2.00"})
        setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  2.50"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  3.00"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  4.00"})
        setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  5.00"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  7.50"}) : setup_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start  10.0"})


        setup_grid.Rows.Add(New String() {"Motor Reset PB Required?"})
        setup_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 1 end"})
        setup_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 2 ended"})
        setup_grid.Rows.Add(New String() {"E-Stop Cable - Total Feet"})
        setup_grid.Rows.Add(New String() {"E-Stop Cable - Turns"})
        setup_grid.Rows.Add(New String() {"E-Stop PB Station"})
        setup_grid.Rows.Add(New String() {"E-Stop PB Station Guarded"})
        setup_grid.Rows.Add(New String() {"Interlock Box"})
        setup_grid.Rows.Add(New String() {"Pressure Switch"})


        setup_grid.Rows.Add(New String() {"Prox Switch"}) : setup_grid.Rows.Add(New String() {"Limit Switch"}) : setup_grid.Rows.Add(New String() {"Photoeye 2 Hole Mnt."})
        setup_grid.Rows.Add(New String() {"Photoeye 18mm Mnt."}) : setup_grid.Rows.Add(New String() {"PhotoEye Emitter Receiver"}) : setup_grid.Rows.Add(New String() {"PhotoEyeDiffuse 100mm range"})
        setup_grid.Rows.Add(New String() {"PhotoEyeDiffuse450mm range"}) : setup_grid.Rows.Add(New String() {"PhotoEyeDiffuse1000mm range"}) : setup_grid.Rows.Add(New String() {"UltraSonic Prox."})
        setup_grid.Rows.Add(New String() {"ADA Photoeye Brackets Do NOT Clr"})

        '------------------ Load pushbuttons table ---------------------------------
        setup_grid.Rows.Add(New String() {"ADA-2CS101"}) : setup_grid.Rows.Add(New String() {"ADA-3CS101"}) : setup_grid.Rows.Add(New String() {"ADA-3CS102"})
        setup_grid.Rows.Add(New String() {"ADA-3CS103"}) : setup_grid.Rows.Add(New String() {"ADA-3CS104"}) : setup_grid.Rows.Add(New String() {"ADA-3CS105"})
        setup_grid.Rows.Add(New String() {"ADA-1CS102"}) : setup_grid.Rows.Add(New String() {"ADA-1CS103"}) : setup_grid.Rows.Add(New String() {"ADA-1CS104"})
        setup_grid.Rows.Add(New String() {"ADA-1CS101"}) : setup_grid.Rows.Add(New String() {"Guarded Pushbutton Colored Inserts"}) : setup_grid.Rows.Add(New String() {"Pushbutton with Colored Inserts"})
        setup_grid.Rows.Add(New String() {"Green Guarded Lighted"}) : setup_grid.Rows.Add(New String() {"Blue Lighted"}) : setup_grid.Rows.Add(New String() {"Red Lighted"})
        setup_grid.Rows.Add(New String() {"Amber Lighted"}) : setup_grid.Rows.Add(New String() {"Selector Switch 2 POS"}) : setup_grid.Rows.Add(New String() {"Selector Switch 3 POS"})
        setup_grid.Rows.Add(New String() {"Selector Switch 3 POS Keyed"})

        '------------ load lights and motion --------------------
        setup_grid.Rows.Add(New String() {"Pilot Lights Amber"}) : setup_grid.Rows.Add(New String() {"Pilot Lights Blue"}) : setup_grid.Rows.Add(New String() {"Pilot Lights Green"})
        setup_grid.Rows.Add(New String() {"Pilot Lights Red"}) : setup_grid.Rows.Add(New String() {"Pilot Lights White"}) : setup_grid.Rows.Add(New String() {"Stack Light Amber"})
        setup_grid.Rows.Add(New String() {"Stack Light Blue"}) : setup_grid.Rows.Add(New String() {"Stack Light Green"}) : setup_grid.Rows.Add(New String() {"Stack Light Red"})
        setup_grid.Rows.Add(New String() {"Stack Light Horn 85 db"}) : setup_grid.Rows.Add(New String() {"Stack Light Horn 100 db 8 tone"}) : setup_grid.Rows.Add(New String() {"MZ Interlock (1 in 1 motion out)"})
        setup_grid.Rows.Add(New String() {"MZ Output (1 motion out)"}) : setup_grid.Rows.Add(New String() {"Solenoid"}) : setup_grid.Rows.Add(New String() {"Double-Ended Solenoid"})

        '--------------- load field devices ----------------------------
        setup_grid.Rows.Add(New String() {"VFD hardwired provided by others"}) : setup_grid.Rows.Add(New String() {"EZLogic PhotoEye Zone"}) : setup_grid.Rows.Add(New String() {"EZLogic Controlled Zone"})
        setup_grid.Rows.Add(New String() {"ZP Accum Photo Eyes (EZLogic)"}) : setup_grid.Rows.Add(New String() {"Controls MDR non reversing"}) : setup_grid.Rows.Add(New String() {"Controls MDR reversing"})
        setup_grid.Rows.Add(New String() {"Release"}) : setup_grid.Rows.Add(New String() {"ZP Accum Power Supply (DC)"}) : setup_grid.Rows.Add(New String() {"Additional amps drawn off of main 24vdc power supply"})
        setup_grid.Rows.Add(New String() {"12mm shaft Encoder"}) : setup_grid.Rows.Add(New String() {"Wheel Encoder"}) : setup_grid.Rows.Add(New String() {"Analog device"})

        '--------------- load bus extenders ----------------------------------
        setup_grid.Rows.Add(New String() {"B&R X2X Bus Extender  Module"}) : setup_grid.Rows.Add(New String() {"B&R X2X 8 Port 16 I/O IP67 Module"}) : setup_grid.Rows.Add(New String() {"SWD 1 Port 2 I/O Module"})
        setup_grid.Rows.Add(New String() {"SWD 2 Port 4 I/O Module"}) : setup_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as in"}) : setup_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as out"})
        setup_grid.Rows.Add(New String() {"SWD 8 Port 8 In/8 Mot Module"}) : setup_grid.Rows.Add(New String() {"SWD 8 Port 16 Mot Module"}) : setup_grid.Rows.Add(New String() {"IO Link Master Module (4 port)IP20"})
        setup_grid.Rows.Add(New String() {"IO Link Master Module (8 port)IP69"}) : setup_grid.Rows.Add(New String() {"IO Link Slave Module 16in (8 port)IP69"}) : setup_grid.Rows.Add(New String() {"IO Link Slave Module 10in6out (8 port)IP69"})

        '-------------------load scanners -------------------------------------
        setup_grid.Rows.Add(New String() {"Scanner - Line (Sick)"}) : setup_grid.Rows.Add(New String() {"Scanner - Omni (Mini-X)"})
        setup_grid.Rows.Add(New String() {"Cognex Bar Code Scan DataMan KIT"}) : setup_grid.Rows.Add(New String() {"Cognex Insite Vision KIT"})

        '------------------- load brake ---------------------------------------
        setup_grid.Rows.Add(New String() {"Clutch/Brake 120VAC Powered"}) : setup_grid.Rows.Add(New String() {"Clutch/Brake 24 VDC Powered"})
        setup_grid.Rows.Add(New String() {"120VAC 0.5kva Power supply"}) : setup_grid.Rows.Add(New String() {"120VAC 2.0kva Power supply"})
        setup_grid.Rows.Add(New String() {"EZLogic Amps needed at 24vdc"}) : setup_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"})
        setup_grid.Rows.Add(New String() {"Total Aux. Amps needed at 24vdc"}) : setup_grid.Rows.Add(New String() {"Non SWD Class2 PWR (Dumb Box)"})
        setup_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"}) : setup_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Drops"})
        setup_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Ext."}) : setup_grid.Rows.Add(New String() {"Add. Motion Amps needed at 24vdc"})
        setup_grid.Rows.Add(New String() {"Amps needed at 24vdc (260 Max)"}) : setup_grid.Rows.Add(New String() {"Non SWD MDR PWR (Dumb Box)"})
        setup_grid.Rows.Add(New String() {"Standard Estop / Breaker Inputs"})

        Call ADA_Setup.Cal_total_Setup()

        'setup

        setup_grid.Rows(1).Cells(1).Value() = ADA_Setup.t_NEMA12
        setup_grid.Rows(2).Cells(1).Value() = ADA_Setup.t_NEMA4X
        setup_grid.Rows(3).Cells(1).Value() = ADA_Setup.t_PLC_Gatewat_PGW
        setup_grid.Rows(4).Cells(1).Value() = ADA_Setup.t_control_logic
        setup_grid.Rows(5).Cells(1).Value() = ADA_Setup.t_compact_logic
        setup_grid.Rows(6).Cells(1).Value() = ADA_Setup.t_PLC_UPS
        setup_grid.Rows(7).Cells(1).Value() = ADA_Setup.t_Amp_Mon
        setup_grid.Rows(8).Cells(1).Value() = ADA_Setup.t_Remote_IO
        setup_grid.Rows(9).Cells(1).Value() = ADA_Setup.t_total_current_group
        setup_grid.Rows(11).Cells(1).Value() = ADA_Setup.t_mdr

        'motors
        For i = 0 To 47
            setup_grid.Rows(i + 12).Cells(1).Value() = ADA_Setup.t_Motor_array(i)
        Next
        'inputs
        For i = 0 To 18
            setup_grid.Rows(i + 60).Cells(1).Value() = ADA_Setup.t_inputs_array(i)
        Next
        'push
        For i = 0 To 18
            setup_grid.Rows(i + 79).Cells(1).Value() = ADA_Setup.t_push_array(i)
        Next
        'lights
        For i = 0 To 14
            setup_grid.Rows(i + 98).Cells(1).Value() = ADA_Setup.t_lights_array(i)
        Next
        'field
        For i = 0 To 11
            setup_grid.Rows(i + 113).Cells(1).Value() = ADA_Setup.t_field_array(i)
        Next
        'bus
        For i = 0 To 11
            setup_grid.Rows(i + 125).Cells(1).Value() = ADA_Setup.t_bus_array(i)
        Next
        'scanner
        For i = 0 To 3
            setup_grid.Rows(i + 137).Cells(1).Value() = ADA_Setup.t_scanner_array(i)
        Next
        'brakes
        For i = 0 To 14
            setup_grid.Rows(i + 141).Cells(1).Value() = ADA_Setup.t_brakes_array(i)
        Next


    End Sub

    Private Sub setup_grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles setup_grid.CellContentClick

    End Sub
End Class