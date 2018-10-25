pageextension 50103 GenJournalExt extends "General Journal"
{
    layout
    {
        addbefore(Description)
        {
            field(LeaseNo; LeaseNo)
            {
                ApplicationArea = All;
            }
        }
        modify("Document Type")
        {
            Visible = false;
        }
        addafter("Account No.")
        {
            field("Shortcut Dimension 1 Code78290"; "Shortcut Dimension 1 Code")
            {
                ApplicationArea = All;
            }
        }
        modify("Gen. Posting Type")
        {
            Visible = false;
        }
        modify("Gen. Bus. Posting Group")
        {
            Visible = false;
        }
        modify("Gen. Prod. Posting Group")
        {
            Visible = false;
        }
        modify("Tax Liable")
        {
            Visible = false;
        }
        modify("Tax Area Code")
        {
            Visible = false;
        }
        modify("Tax Group Code")
        {
            Visible = false;
        }
        modify("Amount (LCY)")
        {
            Visible = false;
        }
        modify("Bal. Account Type")
        {
            Visible = false;
        }
        modify("Bal. Account No.")
        {
            Visible = false;
        }
        modify("Bal. Gen. Posting Type")
        {
            Visible = false;
        }
        modify("Bal. Gen. Bus. Posting Group")
        {
            Visible = false;
        }
        modify("Bal. Gen. Prod. Posting Group")
        {
            Visible = false;
        }
        modify("Deferral Code")
        {
            Visible = false;
        }
        modify(Correction)
        {
            Visible = false;
        }
        addafter(Amount)
        {
            field("Shortcut Dimension 2 Code76830"; "Shortcut Dimension 2 Code")
            {
                ApplicationArea = All;
            }
        }
        modify("ShortcutDimCode[3]")
        {
            Visible = true;
        }
        //moveafter("Shortcut Dimension 2 Code76830";"ShortcutDimCode[3]")
        modify("ShortcutDimCode[4]")
        {
            Visible = true;
        }
        moveafter("ShortcutDimCode[3]"; "ShortcutDimCode[4]")
    }
    actions
    {
        // Add changes to page actions here
        addafter(EditInExcel)
        {
            action(Import_Excel)
            {
                ApplicationArea = All;
                Caption = 'Import Excel';
                Image = ImportExcel;

                trigger OnAction()
                begin
                    ImportGenJnlExcel();
                    Rec_ExcelBuffer.DeleteAll();
                end;

            }
            action(Export_Excel)
            {
                ApplicationArea = All;
                Caption = 'Export Excel';
                Image = ExportToExcel;

                trigger OnAction()
                var
                    GenJnlLine: Record "Gen. Journal Line";
                begin
                    Rec_GenJnl.Copy(Rec);
                    Currpage.SetSelectionFilter(Rec_GenJnl);
                    ExportGenJnlBuffer();


                end;
            }

        }
    }

    var
        Rec_ExcelBuffer: Record "Excel Buffer";
        Rows: Integer;
        Columns: Integer;
        Filename: Text;
        FileMgmt: Codeunit "File Management";
        ExcelFile: File;
        Instr: InStream;
        Sheetname: Text;
        FileUploaded: Boolean;
        RowNo: Integer;
        ColNo: Integer;
        Rec_GenJnl: Record "Gen. Journal Line";

    procedure ImportGenJnlExcel()
    var
    begin
        Rec_ExcelBuffer.DeleteAll();
        Rows := 0;
        Columns := 0;
        FileUploaded := UploadIntoStream('Select File to Upload', '', '', Filename, Instr);

        if Filename <> '' then
            Sheetname := Rec_ExcelBuffer.SelectSheetsNameStream(Instr)
        else
            exit;


        Rec_ExcelBuffer.Reset;
        Rec_ExcelBuffer.OpenBookStream(Instr, Sheetname);
        Rec_ExcelBuffer.ReadSheet();

        Commit();
        Rec_ExcelBuffer.Reset();
        Rec_ExcelBuffer.SetRange("Column No.", 1);
        if Rec_ExcelBuffer.FindFirst() then
            repeat
                Rows := Rows + 1;
            until Rec_ExcelBuffer.Next() = 0;
        //Message(Format(Rows));

        Rec_ExcelBuffer.Reset();
        Rec_ExcelBuffer.SetRange("Row No.", 1);
        if Rec_ExcelBuffer.FindFirst() then
            repeat
                Columns := Columns + 1;
            until Rec_ExcelBuffer.Next() = 0;
        //Message(Format(Columns));
        //Modify or Insert
        for RowNo := 2 to Rows do begin
            Rec_GenJnl.Reset();
            if Rec_GenJnl.Get(GetValueAtIndex(RowNo, 1), GetValueAtIndex(RowNo, 2), GetValueAtIndex(RowNo, 3)) then begin
                Evaluate(Rec_GenJnl."Posting Date", GetValueAtIndex(RowNo, 4));
                Rec_GenJnl.Validate("Posting Date");
                Evaluate(Rec_GenJnl."Document No.", GetValueAtIndex(RowNo, 5));
                Rec_GenJnl.Validate("Document No.");
                Evaluate(Rec_GenJnl."Account Type", GetValueAtIndex(RowNo, 6));
                Rec_GenJnl.Validate("Account Type");
                Evaluate(Rec_GenJnl."Account No.", GetValueAtIndex(RowNo, 7));
                Rec_GenJnl.Validate("Account No.");
                Evaluate(Rec_GenJnl."Shortcut Dimension 1 Code", GetValueAtIndex(RowNo, 8));
                Rec_GenJnl.Validate("Shortcut Dimension 1 Code");
                Evaluate(Rec_GenJnl.LeaseNo, GetValueAtIndex(RowNo, 9));
                Rec_GenJnl.Validate(LeaseNo);
                Evaluate(Rec_GenJnl.Description, GetValueAtIndex(RowNo, 10));
                Rec_GenJnl.Validate(Description);
                Evaluate(Rec_GenJnl.Amount, GetValueAtIndex(RowNo, 11));
                Rec_GenJnl.Validate(Amount);
                Evaluate(Rec_GenJnl."Shortcut Dimension 2 Code", GetValueAtIndex(RowNo, 12));
                Rec_GenJnl.Validate(Rec_GenJnl."Shortcut Dimension 2 Code");
                Rec_GenJnl.Modify(true);

            end
            else begin
                Rec_GenJnl.Init();
                Evaluate(Rec_GenJnl."Journal Template Name", GetValueAtIndex(RowNo, 1));

                Evaluate(Rec_GenJnl."Journal Batch Name", GetValueAtIndex(RowNo, 2));
                Evaluate(Rec_GenJnl."Line No.", GetValueAtIndex(RowNo, 3));
                Evaluate(Rec_GenJnl."Posting Date", GetValueAtIndex(RowNo, 4));
                Evaluate(Rec_GenJnl."Document No.", GetValueAtIndex(RowNo, 5));
                Evaluate(Rec_GenJnl."Account Type", GetValueAtIndex(RowNo, 6));
                Evaluate(Rec_GenJnl."Account No.", GetValueAtIndex(RowNo, 7));
                Rec_GenJnl.Validate("Account No.");
                Evaluate(Rec_GenJnl."Shortcut Dimension 1 Code", GetValueAtIndex(RowNo, 8));
                Evaluate(Rec_GenJnl.LeaseNo, GetValueAtIndex(RowNo, 9));
                Evaluate(Rec_GenJnl.Description, GetValueAtIndex(RowNo, 10));
                Evaluate(Rec_GenJnl.Amount, GetValueAtIndex(RowNo, 11));
                Evaluate(Rec_GenJnl."Shortcut Dimension 2 Code", GetValueAtIndex(RowNo, 12));
                Rec_GenJnl.Validate(Amount);
                Rec_GenJnl.Validate("Posting Date");
                Rec_GenJnl.Validate("Document No.");
                Rec_GenJnl.Validate(LeaseNo);
                Rec_GenJnl.Validate("Shortcut Dimension 1 Code");
                Rec_GenJnl.Validate("Shortcut Dimension 2 Code");
                Rec_GenJnl.Insert();
            end;

        end;
        Message('%1 Rows Imported Successfully!!', Rows - 1);


    end;

    local procedure GetValueAtIndex(RowNo: Integer; ColNo: Integer): Text
    var
    begin
        Rec_ExcelBuffer.Reset();
        IF Rec_ExcelBuffer.Get(RowNo, ColNo) then
            exit(Rec_ExcelBuffer."Cell Value as Text");
    end;

    procedure ExportGenJnlBuffer()
    var
        myInt: Integer;
    begin
        ExportHeaderGenJnl();
        Rec_GenJnl.SetRange("Journal Template Name", 'GENERAL');
        if Rec_GenJnl.FindFirst() then begin
            repeat
                Rec_ExcelBuffer.NewRow();
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Journal Template Name"), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Journal Batch Name"), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Rec_GenJnl."Line No.", false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Number);
                Rec_ExcelBuffer.AddColumn(Rec_GenJnl."Posting Date", false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Date);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Document No."), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Account Type"), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Account No."), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Shortcut Dimension 1 Code"), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl.LeaseNo), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl.Description), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn(Rec_GenJnl.Amount, false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Number);
                Rec_ExcelBuffer.AddColumn(Format(Rec_GenJnl."Shortcut Dimension 2 Code"), false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn('', false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
                Rec_ExcelBuffer.AddColumn('', false, '', false, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);

            until Rec_GenJnl.next = 0;
            Rec_ExcelBuffer.CreateNewBook('General Journal');
            Rec_ExcelBuffer.WriteSheet('General Journal', CompanyName(), UserId());
            Rec_ExcelBuffer.CloseBook();
            Rec_ExcelBuffer.OpenExcel();


        end;
    end;

    local procedure ExportHeaderGenJnl()
    begin
        Rec_ExcelBuffer.Reset();
        Rec_ExcelBuffer.DeleteAll();
        Rec_ExcelBuffer.Init();
        Rec_ExcelBuffer.AddColumn('Journal Template', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Batch Name', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Line No.', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Posting Date', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Document No.', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Account Type', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Account No.', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Dept Code', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Lease No.', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Description', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Amount', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Credit Grade Code', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('Customer Type Code', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);
        Rec_ExcelBuffer.AddColumn('SBU Code', false, '', true, false, false, '', Rec_ExcelBuffer."Cell Type"::Text);

    end;

}