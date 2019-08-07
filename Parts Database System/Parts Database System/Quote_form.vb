Public Class Quote_form

    Public donotcompute As Boolean
    Private Sub Quote_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Initialize QUOTE Spreadsheet alike
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


        Quote_grid.Rows.Add(New String() {"", "Starter Panel – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(0).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Starter Panel - Labor", "", "", "1", ADA_Setup.totals_grid.Rows(1).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "IO Panel – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(2).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "IO Panel – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(3).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "PLC Panel – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(4).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "PLC Panel – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(5).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})

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

        Quote_grid.Rows.Add(New String() {"", "Disconnects", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "VFD IO modules", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Large HP motors (10HP up)", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "100 HP line item", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Electric heater for I/O panel in freezer", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20%", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20%", "", "", "", "Project Materials"}) : Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})

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

        Quote_grid.Rows.Add(New String() {"", "Field Parts – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(7).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Field Parts – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(8).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "Scanners – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(9).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Scanners – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(10).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "M12 Cables – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(11).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "M12 Cables – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(12).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})
        Quote_grid.Rows.Add(New String() {"", "M12 Estop Cables – Materials", "", "", "1", ADA_Setup.totals_grid.Rows(13).Cells(1).Value(), "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "M12 Estop Cables – Labor", "", "", "1", ADA_Setup.totals_grid.Rows(14).Cells(1).Value(), "", "20%", "", "", "", "Manufacturing Labor"})

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

        Quote_grid.Rows.Add(New String() {"", "PC", "", "", "0", "$1800", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Enclosure + Shipping", "", "", "0", "$1000", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Power Supplies / Enclosures", "", "", "0", "$325", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Extra ControlLogix Ethernet card", "", "", "0", "$2000", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Ethernet card adder to VFD", "", "", "0", "$200", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "hand scanner (USB)", "", "", "0", "$200", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "web camera (monitoring)", "", "", "0", "$400", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "standard laser printer", "", "", "0", "$250", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "label printer", "", "", "0", "$1299", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Cognex Technican (per day)", "", "", "0", "$1300", "", "20%", "", "", "", "Project Subcontract"})
        Quote_grid.Rows.Add(New String() {"", "Scanner stand with installation", "", "", "0", "$1300", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Sick bin full PE", "", "", "0", "$500", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "PanelView", "", "", "0", "$3400", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "Additional Scanner Cost", "", "", "0", "$1500", "", "20%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "", "", "", "0", "", "", "20%", "", "", "", "Project Materials"})

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

        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Labor", "1", "", "1", ADA_Setup.totals_grid.Rows(20).Cells(1).Value(), "", "15%", "", "", "", "Install Labor"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Materials", "1", "", "1", ADA_Setup.totals_grid.Rows(21).Cells(1).Value(), "", "15%", "", "", "", "Install Materials"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Expenses", "1", "", "1", ADA_Setup.totals_grid.Rows(22).Cells(1).Value(), "", "0%", "", "", "", "Install Expenses"})
        Quote_grid.Rows.Add(New String() {"", "Electrical Installation Subcontract", "1", "", "1", ADA_Setup.totals_grid.Rows(23).Cells(1).Value(), "", "15%", "", "", "", "Install Subcontract"})


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

        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING I/O", ADA_Setup.totals_grid.Rows(16).Cells(1).Value(), "", "", "$300", "", "0%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING Motor Big", ADA_Setup.totals_grid.Rows(17).Cells(1).Value(), "", "", "$300", "", "0%", "", "", "", "Project Materials"})
        Quote_grid.Rows.Add(New String() {"", "SHIPPING & HANDLING – Scanners", ADA_Setup.totals_grid.Rows(18).Cells(1).Value(), "", "", "$300", "", "0%", "", "", "", "Project Materials"})

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

        Call compute()

    End Sub

    Private Sub Quote_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Quote_grid.CellValueChanged

        Call compute()
    End Sub

    Sub compute()
        If donotcompute = True Then

            For i = 0 To 93
                'calculations from 0 to 93 row

                If (IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value())) Then
                    Quote_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = False, 0, Quote_grid.Rows(i).Cells(4).Value())
                    Quote_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = False, 0, Quote_grid.Rows(i).Cells(5).Value())

                    Quote_grid.Rows(i).Cells(6).Value() = Quote_grid.Rows(i).Cells(4).Value() * Quote_grid.Rows(i).Cells(5).Value()
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(7).Value())) Then
                    Quote_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(5).Value()) = False, 0, Quote_grid.Rows(i).Cells(5).Value())
                    Quote_grid.Rows(i).Cells(7).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(7).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(7).Value()) = False, 0, Quote_grid.Rows(i).Cells(7).Value())

                    Quote_grid.Rows(i).Cells(8).Value() = (Quote_grid.Rows(i).Cells(7).Value() / 100 + 1) + Quote_grid.Rows(i).Cells(5).Value()
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value())) Then
                    Quote_grid.Rows(i).Cells(4).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(4).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(4).Value()) = False, 0, Quote_grid.Rows(i).Cells(4).Value())
                    Quote_grid.Rows(i).Cells(8).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(8).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value()) = False, 0, Quote_grid.Rows(i).Cells(8).Value())

                    Quote_grid.Rows(i).Cells(9).Value() = Quote_grid.Rows(i).Cells(4).Value() * Quote_grid.Rows(i).Cells(8).Value()
                End If

                If (IsNumeric(Quote_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(8).Value())) Then
                    Quote_grid.Rows(i).Cells(6).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(6).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(6).Value()) = False, 0, Quote_grid.Rows(i).Cells(6).Value())
                    Quote_grid.Rows(i).Cells(9).Value() = If(String.IsNullOrEmpty(Quote_grid.Rows(i).Cells(9).Value()) = True Or IsNumeric(Quote_grid.Rows(i).Cells(9).Value()) = False, 0, Quote_grid.Rows(i).Cells(9).Value())

                    Quote_grid.Rows(i).Cells(10).Value() = If(Quote_grid.Rows(i).Cells(9).Value() > 0, Math.Round(100 * ((Quote_grid.Rows(i).Cells(9).Value() - Quote_grid.Rows(i).Cells(6).Value()) / Quote_grid.Rows(i).Cells(9).Value()), 1), 0)
                End If
            Next
        End If


    End Sub
End Class