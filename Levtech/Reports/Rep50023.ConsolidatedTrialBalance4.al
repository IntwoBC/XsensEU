report 50023 "Consolidated Trial Balance 4"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Levtech\Reports\ConsolidatedTrialBalance4custom.rdl';
    //ApplicationArea = Suite;
    Caption = 'Consolidated Trial Balance (4) IT';
    //UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;

    dataset
    {
        dataitem("Business Unit"; "Business Unit")
        {
            DataItemTableView = SORTING(Code) WHERE(Consolidate = CONST(true));

            trigger OnAfterGetRecord()
            begin
                j := j + 1;
                if j > ArrayLen(BusUnitColumn) then
                    Error(Text002, ArrayLen(BusUnitColumn));
                BusUnitColumn[j] := "Business Unit";
            end;

            trigger OnPreDataItem()
            begin
                PageGroupNo := 1;
                NextPageGroupNo := 1;

                j := 0;

                if BUFilter <> '' then
                    SetFilter(Code, BUFilter);
            end;

            trigger OnPostDataItem()
            begin
                if PrintInExcel then
                    MakeExcelInfo();
            end;
        }
        dataitem("G/L Account"; "G/L Account")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Global Dimension 1 Filter", "Global Dimension 2 Filter", "Business Unit Filter";
            column(FORMAT_TODAY_0_4_; Format(Today, 0, 4))
            {
            }
            column(STRSUBSTNO_Text003_PeriodText_; StrSubstNo(Text003, PeriodText))
            {
            }
            column(COMPANYNAME; COMPANYPROPERTY.DisplayName)
            {
            }
            column(USERID; UserId)
            {
            }
            column(InThousands; InThousands)
            {
            }
            column(G_L_Account__TABLECAPTION__________GLFilter; TableCaption + ': ' + GLFilter)
            {
            }
            column(GLFilter; GLFilter)
            {
            }
            column(AmountType; AmountType)
            {
            }
            column(EmptyString; '')
            {
            }
            column(BusUnitColumn_1__Code; BusUnitColumn[1].Code)
            {
            }
            column(BusUnitColumn_2__Code; BusUnitColumn[2].Code)
            {
            }
            column(BusUnitColumn_3__Code; BusUnitColumn[3].Code)
            {
            }
            column(BusUnitColumn_4__Code; BusUnitColumn[4].Code)
            {
            }
            column(ConsolidStartDate; ConsolidStartDate)
            {
            }
            column(ConsolidEndDate; ConsolidEndDate)
            {
            }
            column(NextPageGroupNo; NextPageGroupNo)
            {
            }
            column(PageGroupNo; PageGroupNo)
            {
            }
            column(NewPage; "New Page")
            {
            }
            column(AccountType; Format("Account Type", 0, 2))
            {
            }
            column(NoBlankLines; "No. of Blank Lines")
            {
            }
            column(G_L_Account_No_; "No.")
            {
            }
            column(Consolidated_Trial_Balance__4_Caption; Consolidated_Trial_Balance__4_CaptionLbl)
            {
            }
            column(CurrReport_PAGENOCaption; CurrReport_PAGENOCaptionLbl)
            {
            }
            column(Amounts_are_in_whole_1000sCaption; Amounts_are_in_whole_1000sCaptionLbl)
            {
            }
            column(G_L_Account___No__Caption; FieldCaption("No."))
            {
            }
            column(PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaption; PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaptionLbl)
            {
            }
            column(Amount_1__Amount_2__Amount_3__Amount_4_Caption; Amount_1__Amount_2__Amount_3__Amount_4_CaptionLbl)
            {
            }
            column(EliminationAmountCaption; EliminationAmountCaptionLbl)
            {
            }
            column(Amount_1__Amount_2__Amount_3__Amount_4__EliminationAmountCaption; Amount_1__Amount_2__Amount_3__Amount_4__EliminationAmountCaptionLbl)
            {
            }
            dataitem(BlankLineCounter; "Integer")
            {
                DataItemTableView = SORTING(Number);

                trigger OnPreDataItem()
                begin
                    SetRange(Number, 1, "G/L Account"."No. of Blank Lines");
                end;
            }
            dataitem("Integer"; "Integer")
            {
                DataItemTableView = SORTING(Number) WHERE(Number = CONST(1));
                column(G_L_Account___No__; "G/L Account"."No.")
                {
                }
                column(PADSTR_____G_L_Account__Indentation___2___G_L_Account__Name; PadStr('', "G/L Account".Indentation * 2) + "G/L Account".Name)
                {
                }
                column(Amount_1_; Amount[1])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_2_; Amount[2])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_3_; Amount[3])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_4_; Amount[4])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_1__Amount_2__Amount_3__Amount_4_; Amount[1] + Amount[2] + Amount[3] + Amount[4])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(EliminationAmount; EliminationAmount)
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_1__Amount_2__Amount_3__Amount_4__EliminationAmount; Amount[1] + Amount[2] + Amount[3] + Amount[4] + EliminationAmount)
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(G_L_Account___No___Control30; "G/L Account"."No.")
                {
                }
                column(PADSTR_____G_L_Account__Indentation___2___G_L_Account__Name_Control31; PadStr('', "G/L Account".Indentation * 2) + "G/L Account".Name)
                {
                }
                column(Amount_1__Control32; Amount[1])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_2__Control33; Amount[2])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_3__Control34; Amount[3])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_4__Control35; Amount[4])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_1__Amount_2__Amount_3__Amount_4__Control36; Amount[1] + Amount[2] + Amount[3] + Amount[4])
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(EliminationAmount_Control37; EliminationAmount)
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Amount_1__Amount_2__Amount_3__Amount_4__EliminationAmount_Control38; Amount[1] + Amount[2] + Amount[3] + Amount[4] + EliminationAmount)
                {
                    AutoFormatType = 1;
                    DecimalPlaces = 0 : 0;
                }
                column(Integer_Number; Number)
                {
                }

            }

            trigger OnAfterGetRecord()
            begin
                PageGroupNo := NextPageGroupNo;
                if "New Page" then
                    NextPageGroupNo := PageGroupNo + 1;

                if "G/L Account"."Account Type" = "G/L Account"."Account Type"::Posting then
                    IsBold := false
                else
                    IsBold := true;
                if PrintInExcel then begin
                    ExcelBuf.AddColumn("G/L Account"."No.", FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn("G/L Account".Name, FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                end;


                Clear(TotalArrayAmount);
                for i := 1 to J do begin
                    SetRange("Business Unit Filter", BusUnitColumn[i].Code);
                    if (BusUnitColumn[i]."Starting Date" <> 0D) or (BusUnitColumn[i]."Ending Date" <> 0D) then
                        SetRange("Date Filter", BusUnitColumn[i]."Starting Date", BusUnitColumn[i]."Ending Date")
                    else
                        SetRange("Date Filter", ConsolidStartDate, ConsolidEndDate);

                    if AmountType = AmountType::"Net Change" then begin
                        CalcFields("Net Change");
                        Amount[i] := "Net Change";
                    end else begin
                        CalcFields("Balance at Date");
                        Amount[i] := "Balance at Date";
                    end;
                    if InThousands then
                        Amount[i] := Amount[i] / 1000;
                    if PrintInExcel then begin
                        TotalArrayAmount += Amount[i];
                        ExcelBuf.AddColumn(Amount[i], FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    end;
                end;
                SetRange("Date Filter", ConsolidStartDate, ConsolidEndDate);
                SetRange("Business Unit Filter", '');

                if AmountType = AmountType::"Net Change" then begin
                    CalcFields("Net Change");
                    EliminationAmount := "Net Change";
                end else begin
                    CalcFields("Balance at Date");
                    EliminationAmount := "Balance at Date";
                end;
                if InThousands then
                    EliminationAmount := EliminationAmount / 1000;

                if PrintInExcel then begin
                    ExcelBuf.AddColumn(TotalArrayAmount, FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(EliminationAmount, FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalArrayAmount + EliminationAmount, FALSE, '', IsBold, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.NewRow();
                end;
            end;

            trigger OnPreDataItem()
            begin
                PageGroupNo := 1;
                NextPageGroupNo := 1;

                if j = 0 then
                    CurrReport.Break();
            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    group("Consolidation Period")
                    {
                        Caption = 'Consolidation Period';
                        field(ConsolidStartDate; ConsolidStartDate)
                        {
                            ApplicationArea = Suite;
                            Caption = 'Starting Date';
                            ClosingDates = true;
                            ToolTip = 'Specifies the first date in the period from which posted entries in the consolidated company will be shown.';
                        }
                        field(ConsolidEndDate; ConsolidEndDate)
                        {
                            ApplicationArea = Suite;
                            Caption = 'Ending Date';
                            ClosingDates = true;
                            ToolTip = 'Specifies the end date for the period to process. If a business unit has a different fiscal year than the group, enter the end date for this company in the Business Unit window.';
                        }
                    }
                    field(AmountType; AmountType)
                    {
                        ApplicationArea = Suite;
                        Caption = 'Show';
                        ToolTip = 'Specifies if the selected value is shown in the window.';
                    }
                    field(InThousands; InThousands)
                    {
                        ApplicationArea = Suite;
                        Caption = 'Amounts in whole 1000s';
                        ToolTip = 'Specifies if the amounts in the report are shown in whole 1000s.';
                    }
                    field(PrintInExcel; PrintInExcel)
                    {
                        ApplicationArea = All;
                    }
                }
            }
        }

        trigger OnOpenPage()
        begin
            PrintInExcel := true;
        end;
    }

    labels
    {
    }

    trigger OnPreReport()
    begin
        GLFilter := "G/L Account".GetFilters;
        if ConsolidStartDate = 0D then
            Error(Text000);
        if ConsolidEndDate = 0D then
            Error(Text001);
        "G/L Account".SetRange("Date Filter", ConsolidStartDate, ConsolidEndDate);
        PeriodText := "G/L Account".GetFilter("Date Filter");

        BUFilter := "G/L Account".GetFilter("Business Unit Filter");
    end;



    local procedure MakeExcelInfo()
    begin
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text103), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(CompanyName, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text105), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(Text102), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text104), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Consolidated Trial Balance 4", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn("G/L Account".TableCaption, FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(GLFilter, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn('Period', FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(PeriodText, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);

        if InThousands then begin
            ExcelBuf.NewRow;
            ExcelBuf.AddInfoColumn(Amounts_are_in_whole_1000sCaptionLbl, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        end;

        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text106), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text107), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddInfoColumn(TIME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Time);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    local procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        for i := 1 to J do begin
            ExcelBuf.AddColumn(BusUnitColumn[i].Code, FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        end;

        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Eliminations', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Total Inc. Eliminations', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
    END;

    trigger OnPostReport()
    var
        Outstr: OutStream;
        TempBlb: Codeunit "Temp Blob";
        Instr: InStream;
        FileName: Text;
    begin
        if PrintInExcel then begin
            Clear(TempBlb);
            TempBlb.CreateOutStream(Outstr);
            GetExcelInToSteam(Outstr);
            TempBlb.CreateInStream(Instr);
            FileName := Text102 + DelChr(Format(CurrentDateTime), '=', ':AMPM\/ ') + '.xlsx';
            DownloadFromStream(Instr, '', '', '', FileName);
        end;
    end;

    local procedure CreateExcelbook()
    begin
        ExcelBuf.CreateNewBook(Text101);
        ExcelBuf.WriteSheet(Text102, COMPANYNAME, USERID);
        ExcelBuf.CloseBook();
        ExcelBuf.OpenExcel();
        // Not calling this function. Getting excel data in stream to change the downloaded file name
    end;

    procedure GetExcelInToSteam(var ReportInOutStream: OutStream)
    begin
        ExcelBuf.CreateNewBook(Text101);
        ExcelBuf.WriteSheet(Text102, COMPANYNAME, USERID);
        ExcelBuf.CloseBook();
        ExcelBuf.SaveToStream(ReportInOutStream, true);
    end;

    var
        Text000: Label 'Enter the starting date for the consolidation period.';
        Text001: Label 'Enter the ending date for the consolidation period.';
        Text002: Label 'A maximum of %1 consolidating companies can be included in this report.';
        Text003: Label 'Period: %1';
        BusUnitColumn: array[50] of Record "Business Unit";
        ConsolidStartDate: Date;
        ConsolidEndDate: Date;
        InThousands: Boolean;
        AmountType: Enum "Analysis Amount Type";
        GLFilter: Text;
        EliminationAmount: Decimal;
        PeriodText: Text;
        Amount: array[50] of Decimal;
        i: Integer;
        j: Integer;
        BUFilter: Text;
        PageGroupNo: Integer;
        NextPageGroupNo: Integer;
        Consolidated_Trial_Balance__4_CaptionLbl: Label 'Consolidated Trial Balance (4)';
        CurrReport_PAGENOCaptionLbl: Label 'Page';
        Amounts_are_in_whole_1000sCaptionLbl: Label 'Amounts are in whole 1000s.';
        PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaptionLbl: Label 'Name';
        Amount_1__Amount_2__Amount_3__Amount_4_CaptionLbl: Label 'Total';
        EliminationAmountCaptionLbl: Label 'Eliminations';
        Amount_1__Amount_2__Amount_3__Amount_4__EliminationAmountCaptionLbl: Label 'Total Incl. Eliminations';
        ExcelBuf: Record "Excel Buffer" temporary;
        Text103: Label 'Company Name';
        Text102: Label 'Consolidated Trial Balance';
        Text104: Label 'Report No.';
        Text101: Label 'Data';
        Text105: Label 'Report Name';
        Text106: Label 'User ID';
        Text107: Label 'Date / Time';
        PrintInExcel: Boolean;
        TotalArrayAmount: Decimal;
        IsBold: Boolean;
}

