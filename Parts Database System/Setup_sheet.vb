Public Class Setup_sheet
    Private Sub Setup_sheet_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        sheet_grid.Rows.Add(New String() {"Panel Present BOOL"})  '0
        sheet_grid.Rows.Add(New String() {"NEMA12"})
        sheet_grid.Rows.Add(New String() {"NEMA4X"})
        sheet_grid.Rows.Add(New String() {"PLC in GateWay, PGW"})
        sheet_grid.Rows.Add(New String() {"ControlLogix PLC Box"})
        sheet_grid.Rows.Add(New String() {"CompactLogix PLC Box"})
        sheet_grid.Rows.Add(New String() {"PLC UPS (Battery Backup)"})
        sheet_grid.Rows.Add(New String() {"Amp Mon. Electro Motor Starters"})
        sheet_grid.Rows.Add(New String() {"Remote I/O Only"})
        sheet_grid.Rows.Add(New String() {"Total Current Per Group"})
        sheet_grid.Rows.Add(New String() {"30 Amp Drops"})
        sheet_grid.Rows.Add(New String() {"Motor Driven Rollers (MDRs)"})
        sheet_grid.Rows.Add(New String() {""})
        sheet_grid.Rows.Add(New String() {"Motors Non-Reversing (Horsepower)"})
        sheet_grid.Rows.Add(New String() {"0.25"})
        sheet_grid.Rows.Add(New String() {"0.50"})
        sheet_grid.Rows.Add(New String() {"0.75"})
        sheet_grid.Rows.Add(New String() {"1.00"})
        sheet_grid.Rows.Add(New String() {"1.50"})
        sheet_grid.Rows.Add(New String() {"2.00"})
        sheet_grid.Rows.Add(New String() {"2.50"})
        sheet_grid.Rows.Add(New String() {"3.00"})
        sheet_grid.Rows.Add(New String() {"4.00"})
        sheet_grid.Rows.Add(New String() {"5.00"})
        sheet_grid.Rows.Add(New String() {"7.50"})
        sheet_grid.Rows.Add(New String() {"10.00"})
        sheet_grid.Rows.Add(New String() {"Motors Reversing (Horsepower)"})
        sheet_grid.Rows.Add(New String() {"0.25"})
        sheet_grid.Rows.Add(New String() {"0.50"})
        sheet_grid.Rows.Add(New String() {"0.75"})
        sheet_grid.Rows.Add(New String() {"1.00"})
        sheet_grid.Rows.Add(New String() {"1.50"})
        sheet_grid.Rows.Add(New String() {"2.00"})
        sheet_grid.Rows.Add(New String() {"2.50"})
        sheet_grid.Rows.Add(New String() {"3.00"})
        sheet_grid.Rows.Add(New String() {"4.00"})
        sheet_grid.Rows.Add(New String() {"5.00"})
        sheet_grid.Rows.Add(New String() {"7.50"})
        sheet_grid.Rows.Add(New String() {"10.00"})
        sheet_grid.Rows.Add(New String() {"Full Function VFDs (Horsepower)"})
        sheet_grid.Rows.Add(New String() {"0.25"})
        sheet_grid.Rows.Add(New String() {"0.50"})
        sheet_grid.Rows.Add(New String() {"0.75"})
        sheet_grid.Rows.Add(New String() {"1.00"})
        sheet_grid.Rows.Add(New String() {"1.50"})
        sheet_grid.Rows.Add(New String() {"2.00"})
        sheet_grid.Rows.Add(New String() {"2.50"})
        sheet_grid.Rows.Add(New String() {"3.00"})
        sheet_grid.Rows.Add(New String() {"4.00"})
        sheet_grid.Rows.Add(New String() {"5.00"})
        sheet_grid.Rows.Add(New String() {"7.50"})
        sheet_grid.Rows.Add(New String() {"10.00"})
        sheet_grid.Rows.Add(New String() {"Soft VFDs (H.P.) Prev. Soft Start"})
        sheet_grid.Rows.Add(New String() {"0.25"})
        sheet_grid.Rows.Add(New String() {"0.50"})
        sheet_grid.Rows.Add(New String() {"0.75"})
        sheet_grid.Rows.Add(New String() {"1.00"})
        sheet_grid.Rows.Add(New String() {"1.50"})
        sheet_grid.Rows.Add(New String() {"2.00"})
        sheet_grid.Rows.Add(New String() {"2.50"})
        sheet_grid.Rows.Add(New String() {"3.00"})
        sheet_grid.Rows.Add(New String() {"4.00"})
        sheet_grid.Rows.Add(New String() {"5.00"})
        sheet_grid.Rows.Add(New String() {"7.50"})
        sheet_grid.Rows.Add(New String() {"10.00"})
        sheet_grid.Rows.Add(New String() {"Extra Enclosures with disconnects"})  '65
        sheet_grid.Rows.Add(New String() {"ADA-NEMA12-BOX24(No Starters)"})
        sheet_grid.Rows.Add(New String() {"ADA-NEMA12-BOX30(No Starters)"})
        sheet_grid.Rows.Add(New String() {"Motor Disconnects Suggested for each Amp Monitoring Solid State Motor Starter"})  '68
        sheet_grid.Rows.Add(New String() {"Motor Disconnects"})
        sheet_grid.Rows.Add(New String() {"Motor Reset Input"})  '70
        sheet_grid.Rows.Add(New String() {"Motor Reset PB Required?"})
        sheet_grid.Rows.Add(New String() {"Inputs"}) '72
        sheet_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 1 end"})
        sheet_grid.Rows.Add(New String() {"E-Stop Pull Cord COS 2 ended"})
        sheet_grid.Rows.Add(New String() {"E-Stop Cable - Total Feet"})
        sheet_grid.Rows.Add(New String() {"E-Stop Cable - Turns"})
        sheet_grid.Rows.Add(New String() {"E-Stop PB Station"})
        sheet_grid.Rows.Add(New String() {"E-Stop PB Station Guarded"})
        sheet_grid.Rows.Add(New String() {"Interlock Box"})
        sheet_grid.Rows.Add(New String() {"Pressure Switch"})
        sheet_grid.Rows.Add(New String() {"Prox Switch"})
        sheet_grid.Rows.Add(New String() {"Limit Switch"})
        sheet_grid.Rows.Add(New String() {"Photoeye 2 Hole Mnt."})
        sheet_grid.Rows.Add(New String() {"Photoeye 18mm Mnt."})
        sheet_grid.Rows.Add(New String() {"PhotoEye Emitter Receiver"})
        sheet_grid.Rows.Add(New String() {"PhotoEyeDiffuse 100mm range"})
        sheet_grid.Rows.Add(New String() {"PhotoEyeDiffuse 450mm range"})
        sheet_grid.Rows.Add(New String() {"PhotoEyeDiffuse 1000mm range"})
        sheet_grid.Rows.Add(New String() {"UltraSonic Prox."})
        sheet_grid.Rows.Add(New String() {"ADA Photoeye Brackets"})
        sheet_grid.Rows.Add(New String() {"Pushbuttons"})  '91
        sheet_grid.Rows.Add(New String() {"ADA-ASM-2CS101"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-3CS101"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-3CS102"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-3CS103"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-3CS104"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-3CS105"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-1CS102"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-1CS103"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-1CS104"})
        sheet_grid.Rows.Add(New String() {"ADA-ASM-1CS101"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"Pilot Lights"})  '111
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"User Defined I/O"})
        sheet_grid.Rows.Add(New String() {"Stack Light"})  '117
        sheet_grid.Rows.Add(New String() {"Amber"})
        sheet_grid.Rows.Add(New String() {"Blue"})
        sheet_grid.Rows.Add(New String() {"Green"})
        sheet_grid.Rows.Add(New String() {"Red"})
        sheet_grid.Rows.Add(New String() {"Horn 85 db"})
        sheet_grid.Rows.Add(New String() {"Horn 100 db 8 tone"})
        sheet_grid.Rows.Add(New String() {"Motion Outputs"})    '124
        sheet_grid.Rows.Add(New String() {"MZ Interlock (1 in 1 motion out)"})
        sheet_grid.Rows.Add(New String() {"MZ Output (1 motion out)"})
        sheet_grid.Rows.Add(New String() {"Solenoid"})
        sheet_grid.Rows.Add(New String() {"Double-Ended Solenoid"})
        sheet_grid.Rows.Add(New String() {"Field Devices"})  '129
        sheet_grid.Rows.Add(New String() {"VFD hardwired provided by others"})
        sheet_grid.Rows.Add(New String() {"EZLogic PhotoEye Zone"})
        sheet_grid.Rows.Add(New String() {"EZLogic Controlled Zone"})
        sheet_grid.Rows.Add(New String() {"ZP Accum Photo Eyes (EZLogic)"})
        sheet_grid.Rows.Add(New String() {"Controls MDR non reversing"})
        sheet_grid.Rows.Add(New String() {"Controls MDR reversing"})
        sheet_grid.Rows.Add(New String() {"Release"})
        sheet_grid.Rows.Add(New String() {"ZP Accum Power Supply (DC)"})
        sheet_grid.Rows.Add(New String() {"Additional amps drawn off of main 24vdc power supply"})
        sheet_grid.Rows.Add(New String() {"12mm shaft Encoder"})
        sheet_grid.Rows.Add(New String() {"Wheel Encoder"})
        sheet_grid.Rows.Add(New String() {"Analog device"})
        sheet_grid.Rows.Add(New String() {"Remote I/O & Bus Extenders"})  '142
        sheet_grid.Rows.Add(New String() {"B&R X2X Bus Extender  Module"})
        sheet_grid.Rows.Add(New String() {"B&R X2X 8 Port 16 I/O IP67 Module"})
        sheet_grid.Rows.Add(New String() {"SWD 1 Port 2 I/O Module"})
        sheet_grid.Rows.Add(New String() {"SWD 2 Port 4 I/O Module"})
        sheet_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as in"})
        sheet_grid.Rows.Add(New String() {"SWD 8 Port 16 I/O Module as out"})
        sheet_grid.Rows.Add(New String() {"SWD 8 Port 8 In/8 Mot Module"})
        sheet_grid.Rows.Add(New String() {"SWD 8 Port 16 Mot Module"})
        sheet_grid.Rows.Add(New String() {"IO Link Master Module (4 port)IP20"})
        sheet_grid.Rows.Add(New String() {"IO Link Master Module (8 port)IP69"})
        sheet_grid.Rows.Add(New String() {"IO Link Slave Module 16in (8 port)IP69"})
        sheet_grid.Rows.Add(New String() {"IO Link Slave Module 10in6out (8 port)IP69"})
        sheet_grid.Rows.Add(New String() {"Scanner"})  '155
        sheet_grid.Rows.Add(New String() {"Scanner - Line (Sick)"})
        sheet_grid.Rows.Add(New String() {"Scanner - Omni (Mini-X)"})
        sheet_grid.Rows.Add(New String() {"Cognex Bar Code Scan DataMan KIT"})
        sheet_grid.Rows.Add(New String() {"Cognex Insite Vision KIT"})
        sheet_grid.Rows.Add(New String() {"Clutch/Brake"})  '160
        sheet_grid.Rows.Add(New String() {"120VAC Powered"})
        sheet_grid.Rows.Add(New String() {"24 VDC Powered"})
        sheet_grid.Rows.Add(New String() {"120vac Power Supply"})
        sheet_grid.Rows.Add(New String() {"120VAC 0.5kva Power"})
        sheet_grid.Rows.Add(New String() {"120VAC 2.0kva Power"})
        sheet_grid.Rows.Add(New String() {"Main 24vdc Power Supply"})
        sheet_grid.Rows.Add(New String() {"EZLogic Amps needed at 24vdc"})
        sheet_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"})
        sheet_grid.Rows.Add(New String() {"Total Aux. Amps needed at 24vdc"})
        sheet_grid.Rows.Add(New String() {"Non SWD Class2 PWR (Dumb Box)"})
        sheet_grid.Rows.Add(New String() {"Motion 24vdc Power Supply"})
        sheet_grid.Rows.Add(New String() {"Additional Amps needed at 24vdc"})
        sheet_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Drops"})
        sheet_grid.Rows.Add(New String() {"E24 20 Amp MDR 20' Ext."})
        sheet_grid.Rows.Add(New String() {"Add. Motion Amps needed at 24vdc"})
        sheet_grid.Rows.Add(New String() {"Amps needed at 24vdc (260 Max)"})
        sheet_grid.Rows.Add(New String() {"Non SWD MDR PWR (Dumb Box)"})
        sheet_grid.Rows.Add(New String() {"Standard Estop / Breaker Inputs"})

        'color rows
        sheet_grid.Rows(11).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(13).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(26).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(39).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(52).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(65).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)

        sheet_grid.Rows(68).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(70).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(72).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(91).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(111).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)

        sheet_grid.Rows(117).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(124).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(129).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(142).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(155).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)
        sheet_grid.Rows(160).DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204)

        'create set


        Dim i As Integer : i = 1
        For Each ada_set In ADA_Setup.ADA_set_list

            sheet_grid.Columns.Add(ada_set.myname, ada_set.myname)
            sheet_grid.Columns(i + 3).HeaderCell.Style.BackColor = Color.LightBlue
            sheet_grid.Columns(i + 3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            i = i + 1
        Next


        For i = 4 To sheet_grid.Columns.Count - 1
            sheet_grid.Rows(9).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(10).Cells(i).Style.BackColor = Color.Orange

            sheet_grid.Rows(9).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(10).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(102).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(103).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(104).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(105).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(106).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(107).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(108).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(109).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(110).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(112).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(113).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(114).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(115).Cells(i).Style.BackColor = Color.Orange
            sheet_grid.Rows(116).Cells(i).Style.BackColor = Color.Orange

        Next

        Dim j As Integer : j = 4
        For Each ada_set In ADA_Setup.ADA_set_list

            sheet_grid.Rows(0).Cells(j).Value = ada_set.Panel_bool
            sheet_grid.Rows(1).Cells(j).Value = ada_set.NEMA12
            sheet_grid.Rows(2).Cells(j).Value = ada_set.NEMA4X
            sheet_grid.Rows(3).Cells(j).Value = ada_set.PLC_GateWay_PGW
            sheet_grid.Rows(4).Cells(j).Value = ada_set.ControlLogix_PLC_Box
            sheet_grid.Rows(5).Cells(j).Value = ada_set.CompactLogix_PLC_Box
            sheet_grid.Rows(6).Cells(j).Value = ada_set.PLC_UPS_Battery_Backup
            sheet_grid.Rows(7).Cells(j).Value = ada_set.Amp_Mon_Electro
            sheet_grid.Rows(8).Cells(j).Value = ada_set.Remote_IO
            sheet_grid.Rows(9).Cells(j).Value = ada_set.Total_current_per_group
            sheet_grid.Rows(10).Cells(j).Value = ada_set.Amp_30_Drops

            sheet_grid.Rows(12).Cells(j).Value = If(String.Equals(ada_set.MDRs, "") = True, "", ada_set.MDRs)

            For z = 0 To 11
                If String.Equals("0", ada_set.Motors_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 14).Cells(j).Value = ada_set.Motors_table(z, 0)
                End If
            Next

            For z = 12 To 23
                If String.Equals("0", ada_set.Motors_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 15).Cells(j).Value = ada_set.Motors_table(z, 0)
                End If
            Next

            For z = 24 To 35
                If String.Equals("0", ada_set.Motors_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 16).Cells(j).Value = ada_set.Motors_table(z, 0)
                End If
            Next

            For z = 36 To 47
                If String.Equals("0", ada_set.Motors_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 17).Cells(j).Value = ada_set.Motors_table(z, 0)
                End If
            Next

            sheet_grid.Rows(66).Cells(j).Value = If(String.Equals(ada_set.ADA_NEMA12_BOX24_No_Starters, "") = True, "", ada_set.ADA_NEMA12_BOX24_No_Starters)
            sheet_grid.Rows(67).Cells(j).Value = If(String.Equals(ada_set.ADA_NEMA12_BOX30_No_Starters, "") = True, "", ada_set.ADA_NEMA12_BOX30_No_Starters)
            sheet_grid.Rows(69).Cells(j).Value = If(String.Equals(ada_set.motor_disconnect, "") = True, "", ada_set.motor_disconnect)

            sheet_grid.Rows(71).Cells(j).Value = If(String.Equals(ada_set.inputs_table(0, 0), "0") = True, "", ada_set.inputs_table(0, 0))
            sheet_grid.Rows(71).Cells(1).Value = If(String.Equals(ada_set.inputs_table(0, 1), "0") = True, "", ada_set.inputs_table(0, 1))
            sheet_grid.Rows(71).Cells(2).Value = If(String.Equals(ada_set.inputs_table(0, 2), "0") = True, "", ada_set.inputs_table(0, 2))
            sheet_grid.Rows(71).Cells(3).Value = If(String.Equals(ada_set.inputs_table(0, 3), "0") = True, "", ada_set.inputs_table(0, 3))

            'Inputs
            For z = 1 To 18
                If String.Equals("0", ada_set.inputs_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 72).Cells(j).Value = ada_set.inputs_table(z, 0)
                End If

                sheet_grid.Rows(z + 72).Cells(1).Value = If(String.Equals(ada_set.inputs_table(z, 1), "0") = True, "", ada_set.inputs_table(z, 1))
                sheet_grid.Rows(z + 72).Cells(2).Value = If(String.Equals(ada_set.inputs_table(z, 2), "0") = True, "", ada_set.inputs_table(z, 2))
                sheet_grid.Rows(z + 72).Cells(3).Value = If(String.Equals(ada_set.inputs_table(z, 3), "0") = True, "", ada_set.inputs_table(z, 3))
            Next


            'pushbutton
            For z = 0 To 18
                If String.Equals("0", ada_set.push_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 92).Cells(j).Value = ada_set.push_table(z, 0)
                End If

                sheet_grid.Rows(z + 92).Cells(1).Value = If(String.Equals(ada_set.push_table(z, 1), "0") = True, "", ada_set.push_table(z, 1))
                sheet_grid.Rows(z + 92).Cells(2).Value = If(String.Equals(ada_set.push_table(z, 2), "0") = True, "", ada_set.push_table(z, 2))
                sheet_grid.Rows(z + 92).Cells(3).Value = If(String.Equals(ada_set.push_table(z, 3), "0") = True, "", ada_set.push_table(z, 3))
            Next

            'lights and motion
            For z = 0 To 4
                sheet_grid.Rows(z + 112).Cells(1).Value = If(String.Equals(ada_set.lights_table(z, 1), "0") = True, "", ada_set.lights_table(z, 1))
                sheet_grid.Rows(z + 112).Cells(2).Value = If(String.Equals(ada_set.lights_table(z, 2), "0") = True, "", ada_set.lights_table(z, 2))
                sheet_grid.Rows(z + 112).Cells(3).Value = If(String.Equals(ada_set.lights_table(z, 3), "0") = True, "", ada_set.lights_table(z, 3))
            Next

            For z = 5 To 10
                If String.Equals("0", ada_set.lights_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 118).Cells(j).Value = ada_set.lights_table(z, 0)
                End If
                sheet_grid.Rows(z + 118).Cells(1).Value = If(String.Equals(ada_set.lights_table(z, 1), "0") = True, "", ada_set.lights_table(z, 1))
                sheet_grid.Rows(z + 118).Cells(2).Value = If(String.Equals(ada_set.lights_table(z, 2), "0") = True, "", ada_set.lights_table(z, 2))
                sheet_grid.Rows(z + 118).Cells(3).Value = If(String.Equals(ada_set.lights_table(z, 3), "0") = True, "", ada_set.lights_table(z, 3))
            Next

            'motion
            For z = 11 To 14
                If String.Equals("0", ada_set.lights_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 125).Cells(j).Value = ada_set.lights_table(z, 0)
                End If
                sheet_grid.Rows(z + 125).Cells(1).Value = If(String.Equals(ada_set.lights_table(z, 1), "0") = True, "", ada_set.lights_table(z, 1))
                sheet_grid.Rows(z + 125).Cells(2).Value = If(String.Equals(ada_set.lights_table(z, 2), "0") = True, "", ada_set.lights_table(z, 2))
                sheet_grid.Rows(z + 125).Cells(3).Value = If(String.Equals(ada_set.lights_table(z, 3), "0") = True, "", ada_set.lights_table(z, 3))
            Next


            'field
            For z = 0 To 11
                If String.Equals("0", ada_set.field_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 130).Cells(j).Value = ada_set.field_table(z, 0)
                End If

                sheet_grid.Rows(z + 130).Cells(1).Value = If(String.Equals(ada_set.field_table(z, 1), "0") = True, "", ada_set.field_table(z, 1))
                sheet_grid.Rows(z + 130).Cells(2).Value = If(String.Equals(ada_set.field_table(z, 2), "0") = True, "", ada_set.field_table(z, 2))
                sheet_grid.Rows(z + 130).Cells(3).Value = If(String.Equals(ada_set.field_table(z, 3), "0") = True, "", ada_set.field_table(z, 3))
            Next

            'bus
            For z = 0 To 11
                If String.Equals("0", ada_set.bus_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 143).Cells(j).Value = ada_set.bus_table(z, 0)
                End If

                sheet_grid.Rows(z + 143).Cells(1).Value = If(String.Equals(ada_set.bus_table(z, 1), "0") = True, "", ada_set.bus_table(z, 1))
                sheet_grid.Rows(z + 143).Cells(2).Value = If(String.Equals(ada_set.bus_table(z, 2), "0") = True, "", ada_set.bus_table(z, 2))
                sheet_grid.Rows(z + 143).Cells(3).Value = If(String.Equals(ada_set.bus_table(z, 3), "0") = True, "", ada_set.bus_table(z, 3))
            Next

            'scanner
            For z = 0 To 3
                If String.Equals("0", ada_set.scanner_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 156).Cells(j).Value = ada_set.scanner_table(z, 0)
                End If

                sheet_grid.Rows(z + 156).Cells(1).Value = If(String.Equals(ada_set.scanner_table(z, 1), "0") = True, "", ada_set.scanner_table(z, 1))
                sheet_grid.Rows(z + 156).Cells(2).Value = If(String.Equals(ada_set.scanner_table(z, 2), "0") = True, "", ada_set.scanner_table(z, 2))
                sheet_grid.Rows(z + 156).Cells(3).Value = If(String.Equals(ada_set.scanner_table(z, 3), "0") = True, "", ada_set.scanner_table(z, 3))
            Next

            'brakes
            For z = 0 To 1
                If String.Equals("0", ada_set.brakes_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 161).Cells(j).Value = ada_set.brakes_table(z, 0)
                End If

                sheet_grid.Rows(z + 161).Cells(1).Value = If(String.Equals(ada_set.brakes_table(z, 1), "0") = True, "", ada_set.brakes_table(z, 1))
                sheet_grid.Rows(z + 161).Cells(2).Value = If(String.Equals(ada_set.brakes_table(z, 2), "0") = True, "", ada_set.brakes_table(z, 2))
                sheet_grid.Rows(z + 161).Cells(3).Value = If(String.Equals(ada_set.brakes_table(z, 3), "0") = True, "", ada_set.brakes_table(z, 3))
            Next

            For z = 2 To 3
                If String.Equals("0", ada_set.brakes_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 162).Cells(j).Value = ada_set.brakes_table(z, 0)
                End If

                sheet_grid.Rows(z + 162).Cells(1).Value = If(String.Equals(ada_set.brakes_table(z, 1), "0") = True, "", ada_set.brakes_table(z, 1))
                sheet_grid.Rows(z + 162).Cells(2).Value = If(String.Equals(ada_set.brakes_table(z, 2), "0") = True, "", ada_set.brakes_table(z, 2))
                sheet_grid.Rows(z + 162).Cells(3).Value = If(String.Equals(ada_set.brakes_table(z, 3), "0") = True, "", ada_set.brakes_table(z, 3))
            Next

            For z = 4 To 7
                If String.Equals("0", ada_set.brakes_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 163).Cells(j).Value = ada_set.brakes_table(z, 0)
                End If

                sheet_grid.Rows(z + 163).Cells(1).Value = If(String.Equals(ada_set.brakes_table(z, 1), "0") = True, "", ada_set.brakes_table(z, 1))
                sheet_grid.Rows(z + 163).Cells(2).Value = If(String.Equals(ada_set.brakes_table(z, 2), "0") = True, "", ada_set.brakes_table(z, 2))
                sheet_grid.Rows(z + 163).Cells(3).Value = If(String.Equals(ada_set.brakes_table(z, 3), "0") = True, "", ada_set.brakes_table(z, 3))
            Next


            For z = 8 To 14
                If String.Equals("0", ada_set.brakes_table(z, 0)) = False Then
                    sheet_grid.Rows(z + 164).Cells(j).Value = ada_set.brakes_table(z, 0)
                End If

                sheet_grid.Rows(z + 164).Cells(1).Value = If(String.Equals(ada_set.brakes_table(z, 1), "0") = True, "", ada_set.brakes_table(z, 1))
                sheet_grid.Rows(z + 164).Cells(2).Value = If(String.Equals(ada_set.brakes_table(z, 2), "0") = True, "", ada_set.brakes_table(z, 2))
                sheet_grid.Rows(z + 164).Cells(3).Value = If(String.Equals(ada_set.brakes_table(z, 3), "0") = True, "", ada_set.brakes_table(z, 3))
            Next

            j = j + 1
        Next

    End Sub
End Class