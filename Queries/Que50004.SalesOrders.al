query 50004 SalesOrders
{
    Caption = 'SalesOrders';
    QueryType = Normal;


    elements
    {
        dataitem(Sales_Header; "Sales Header")
        {
            DataItemTableFilter = "Document Type" = const(Order);
            column(Document_Type; "Document Type")
            {

            }
            column(No; "No.")
            {
            }
            column(Posting_Date; "Posting Date")
            {

            }
            column(Sell_to_Customer_No; "Sell-to Customer No.")
            {
            }
            column(Sell_to_Customer_Name; "Sell-to Customer Name")
            {
            }
            column(Your_Reference; "Your Reference")
            {
            }
            column(Order_Date; "Order Date")
            {
            }
            column(Order_Confirmation_Date; "Date Order confimation")
            {
                Caption = 'Order_Confirmation_Date';
            }
            column(Sell_to_Country_Region_Code; "Sell-to Country/Region Code")
            {
            }
            column(Customer_Posting_Group; "Customer Posting Group")
            {
            }
            column(Due_Date; "Due Date")
            {
            }
            column(AllowLineDisc_Hdr; "Allow Line Disc.")
            {
            }
            column(Amount_Hdr; Amount)
            {
            }
            column(AmountIncludingVAT_Hdr; "Amount Including VAT")
            {
            }
            column(AmtShipNotInvLCY; "Amt. Ship. Not Inv. (LCY)")
            {
            }
            column(AmtShipNotInvLCYBase; "Amt. Ship. Not Inv. (LCY) Base")
            {
            }
            column(Application; Application)
            {
            }
            column(Applicationarea; "Application area")
            {
            }
            column(AppliestoDocNo; "Applies-to Doc. No.")
            {
            }
            column(AppliestoDocType; "Applies-to Doc. Type")
            {
            }
            column(AppliestoID; "Applies-to ID")
            {
            }
            column(Archiveoption; "Archive option")
            {
            }
            column("Area_Hdr"; "Area")
            {
            }
            column(AssignedUserID; "Assigned User ID")
            {
            }
            column(BalAccountNo; "Bal. Account No.")
            {
            }
            column(BalAccountType; "Bal. Account Type")
            {
            }
            column(BankAccountCode; "Bank Account Code")
            {
            }
            column(BilltoAddress; "Bill-to Address")
            {
            }
            column(BilltoAddress2; "Bill-to Address 2")
            {
            }
            column(BilltoCity; "Bill-to City")
            {
            }
            column(BilltoContact; "Bill-to Contact")
            {
            }
            column(BilltoContactNo; "Bill-to Contact No.")
            {
            }
            column(BilltoCountryRegionCode; "Bill-to Country/Region Code")
            {
            }
            column(BilltoCounty; "Bill-to County")
            {
            }
            column(BilltoCustomerNo_Hdr; "Bill-to Customer No.")
            {
            }
            column(BilltoCustomerTemplCode; "Bill-to Customer Templ. Code")
            {
            }
            column(BilltoEmail; "Bill-to Email")
            {
            }
            column(BilltoICPartnerCode; "Bill-to IC Partner Code")
            {
            }
            column(BilltoName; "Bill-to Name")
            {
            }
            column(BilltoName2; "Bill-to Name 2")
            {
            }
            column(BilltoPostCode; "Bill-to Post Code")
            {
            }
            column(Boekingsdatumakkoord; "Boekingsdatum akkoord")
            {
            }
            column(CampaignNo; "Campaign No.")
            {
            }
            column(CombineShipments; "Combine Shipments")
            {
            }
            column(Comment; Comment)
            {
            }
            column(Comment2; "Comment 2")
            {
            }
            column(CommentExternal; "Comment External")
            {
            }
            column(CommentInternal; "Comment Internal")
            {
            }
            column(CompanyBankAccountCode; "Company Bank Account Code")
            {
            }
            column(CompletelyShipped_Hdr; "Completely Shipped")
            {
            }
            column(CompressPrepayment; "Compress Prepayment")
            {
            }
            column(Correction; Correction)
            {
            }
            column(CreatedByRapidi; "Created By Rapidi")
            {
            }
            column(CurrencyCode_Hdr; "Currency Code")
            {
            }
            column(CurrencyCodeIT; "Currency Code IT")
            {
            }
            column(CurrencyFactor; "Currency Factor")
            {
            }
            column(CustomerDiscGroup_Hdr; "Customer Disc. Group")
            {
            }
            column(CustomerPriceGroup_Hdr; "Customer Price Group")
            {
            }
            column(DeclarationStatement; "Declaration Statement")
            {
            }
            column(DimensionSetID_Hdr; "Dimension Set ID")
            {
            }
            column(DirectDebitMandateID; "Direct Debit Mandate ID")
            {
            }
            column(DocNoOccurrence; "Doc. No. Occurrence")
            {
            }
            column(EU3PartyTrade; "EU 3-Party Trade")
            {
            }
            column(ExitPoint_Hdr; "Exit Point")
            {
            }
            column(ExternalDocumentNo; "External Document No.")
            {
            }
            column(ExternalID_Hdr; ExternalID)
            {
            }
            column(GenBusPostingGroup_Hdr; "Gen. Bus. Posting Group")
            {
            }
            column(GetShipmentUsed; "Get Shipment Used")
            {
            }
            column(ICDirection; "IC Direction")
            {
            }
            column(ICPurchaseorder; "IC Purchase order")
            {
            }
            column(ICPurchaseordercreated; "IC Purchase order created")
            {
            }
            column(ICPurchaseordersend; "IC Purchase order send")
            {
            }
            column(ICStatus; "IC Status")
            {
            }
            column(IncomingDocumentEntryNo; "Incoming Document Entry No.")
            {
            }
            column(Invoice; Invoice)
            {
            }
            column(InvoiceDiscCode; "Invoice Disc. Code")
            {
            }
            column(InvoiceDiscountAmount; "Invoice Discount Amount")
            {
            }
            column(InvoiceDiscountCalculation; "Invoice Discount Calculation")
            {
            }
            column(InvoiceDiscountValue; "Invoice Discount Value")
            {
            }
            column(IsTest; IsTest)
            {
            }
            column(JobQueueEntryID; "Job Queue Entry ID")
            {
            }
            column(JobQueueStatus; "Job Queue Status")
            {
            }
            column(JournalTemplName; "Journal Templ. Name")
            {
            }
            column(LanguageCode; "Language Code")
            {
            }
            column(LastEmailNotifCleared; "Last Email Notif Cleared")
            {
            }
            column(LastEmailSentStatus; "Last Email Sent Status")
            {
            }
            column(LastEmailSentTime; "Last Email Sent Time")
            {
            }
            column(LastPostingNo; "Last Posting No.")
            {
            }
            column(LastPrepaymentNo; "Last Prepayment No.")
            {
            }
            column(LastPrepmtCrMemoNo; "Last Prepmt. Cr. Memo No.")
            {
            }
            column(LastReturnReceiptNo; "Last Return Receipt No.")
            {
            }
            column(LastShipmentDate; "Last Shipment Date")
            {
            }
            column(LastShippingNo; "Last Shipping No.")
            {
            }
            column(LateOrderShipping; "Late Order Shipping")
            {
            }
            column(LocationCode_Hdr; "Location Code")
            {
            }
            column(NoofArchivedVersions; "No. of Archived Versions")
            {
            }
            column(NoPrinted; "No. Printed")
            {
            }
            column(NoSeries; "No. Series")
            {
            }
            column(OnHold; "On Hold")
            {
            }
            column(OpportunityNo; "Opportunity No.")
            {
            }
            column(OrderClass; "Order Class")
            {
            }
            column(OutboundWhseHandlingTime_Hdr; "Outbound Whse. Handling Time")
            {
            }
            column(PackageTrackingNo; "Package Tracking No.")
            {
            }
            column(PaymentDiscount; "Payment Discount %")
            {
            }
            column(PaymentMethodCode; "Payment Method Code")
            {
            }
            column(PaymentMethodCodeIT; "Payment Method Code IT")
            {
            }
            column(PaymentServiceSetID; "Payment Service Set ID")
            {
            }
            column(PaymentTermsCode; "Payment Terms Code")
            {
            }
            column(PaymentTermsCodeIT; "Payment Terms Code IT")
            {
            }
            column(PmtDiscountDate; "Pmt. Discount Date")
            {
            }
            column(PostingDescription; "Posting Description")
            {
            }
            column(PostingfromWhseRef; "Posting from Whse. Ref.")
            {
            }
            column(PostingNo; "Posting No.")
            {
            }
            column(PostingNoSeries; "Posting No. Series")
            {
            }
            column(Prepayment_Hdr; "Prepayment %")
            {
            }
            column(PrepaymentDueDate; "Prepayment Due Date")
            {
            }
            column(PrepaymentNo; "Prepayment No.")
            {
            }
            column(PrepaymentNoSeries; "Prepayment No. Series")
            {
            }
            column(PrepmtCrMemoNo; "Prepmt. Cr. Memo No.")
            {
            }
            column(PrepmtCrMemoNoSeries; "Prepmt. Cr. Memo No. Series")
            {
            }
            column(PrepmtPaymentDiscount; "Prepmt. Payment Discount %")
            {
            }
            column(PrepmtPaymentTermsCode; "Prepmt. Payment Terms Code")
            {
            }
            column(PrepmtPmtDiscountDate; "Prepmt. Pmt. Discount Date")
            {
            }
            column(PrepmtPostingDescription; "Prepmt. Posting Description")
            {
            }
            column(PriceCalculationMethod_Hdr; "Price Calculation Method")
            {
            }
            column(PricesIncludingVAT; "Prices Including VAT")
            {
            }
            column(PrintPostedDocuments; "Print Posted Documents")
            {
            }
            column(PromisedDeliveryDate_Hdr; "Promised Delivery Date")
            {
            }
            column(QuoteAccepted; "Quote Accepted")
            {
            }
            column(QuoteAcceptedDate; "Quote Accepted Date")
            {
            }
            column(QuoteNo; "Quote No.")
            {
            }
            column(QuoteSenttoCustomer; "Quote Sent to Customer")
            {
            }
            column(QuoteValidUntilDate; "Quote Valid Until Date")
            {
            }
            column(RapidiFieldsUpdated; "Rapidi Fields Updated")
            {
            }
            column(ReadytoShip; "Ready to Ship")
            {
            }
            column(ReasonCode; "Reason Code")
            {
            }
            column(RecalculateInvoiceDisc_Hdr; "Recalculate Invoice Disc.")
            {
            }
            column(Receive; Receive)
            {
            }
            column(RequestedDeliveryDate_Hdr; "Requested Delivery Date")
            {
            }
            column(Reserve_Hdr; Reserve)
            {
            }
            column(ResponsibilityCenter_Hdr; "Responsibility Center")
            {
            }
            column(ReturnReceiptNo_Hdr; "Return Receipt No.")
            {
            }
            column(ReturnReceiptNoSeries; "Return Receipt No. Series")
            {
            }
            column(SalesDate; "Sales Date")
            {
            }
            column(SalesForceComment; "SalesForce Comment")
            {
            }
            column(SalespersonCodeIT; "Salesperson Code IT")
            {
            }
            column(Segment; Segment)
            {
            }
            column(SelltoAddress; "Sell-to Address")
            {
            }
            column(SelltoAddress2; "Sell-to Address 2")
            {
            }
            column(SelltoCity; "Sell-to City")
            {
            }
            column(SelltoContact; "Sell-to Contact")
            {
            }
            column(SelltoContactEmail; "Sell-to Contact E-mail")
            {
            }
            column(SelltoContactNo; "Sell-to Contact No.")
            {
            }
            column(SelltoContactPhone; "Sell-to Contact Phone")
            {
            }
            column(SelltoCounty; "Sell-to County")
            {
            }
            column(SelltoCustomerName2; "Sell-to Customer Name 2")
            {
            }
            column(SelltoCustomerName3; "Sell-to Customer Name 3")
            {
            }
            column(SelltoCustomerTemplCode; "Sell-to Customer Templ. Code")
            {
            }
            column(SelltoEMail; "Sell-to E-Mail")
            {
            }
            column(SelltoICAddress; "Sell-to IC Address")
            {
            }
            column(SelltoICAddress2; "Sell-to IC Address 2")
            {
            }
            column(SelltoICCity; "Sell-to IC City")
            {
            }
            column(SelltoICContact; "Sell-to IC Contact")
            {
            }
            column(SelltoICCustomerNo; "Sell-to IC Customer No.")
            {
            }
            column(SelltoICName; "Sell-to IC Name")
            {
            }
            column(SelltoICName2; "Sell-to IC Name 2")
            {
            }
            column(SelltoICPartnerCode; "Sell-to IC Partner Code")
            {
            }
            column(SelltoICPostCode; "Sell-to IC Post Code")
            {
            }
            column(SelltoPhoneNo; "Sell-to Phone No.")
            {
            }
            column(SelltoPostCode; "Sell-to Post Code")
            {
            }
            column(SendICDocument; "Send IC Document")
            {
            }
            column(SentasEmail; "Sent as Email")
            {
            }
            column(Ship; Ship)
            {
            }
            column(ShipmentDate_Hdr; "Shipment Date")
            {
            }
            column(ShipmentDateIT; "Shipment Date IT")
            {
            }
            column(ShipmentMethodCode; "Shipment Method Code")
            {
            }
            column(ShipmentMethodDescription; "Shipment Method Description")
            {
            }
            column(Shipped; Shipped)
            {
            }
            column(ShippedNotInvoiced_Hdr; "Shipped Not Invoiced")
            {
            }
            column(ShippingAdvice; "Shipping Advice")
            {
            }
            column(ShippingAgentCode_Hdr; "Shipping Agent Code")
            {
            }
            column(ShippingAgentServiceCode_Hdr; "Shipping Agent Service Code")
            {
            }
            column(ShippingNo; "Shipping No.")
            {
            }
            column(ShippingNoSeries; "Shipping No. Series")
            {
            }
            column(ShippingTime_Hdr; "Shipping Time")
            {
            }
            column(ShiptoAddress; "Ship-to Address")
            {
            }
            column(ShiptoAddress2; "Ship-to Address 2")
            {
            }
            column(ShiptoCity; "Ship-to City")
            {
            }
            column(ShiptoCode; "Ship-to Code")
            {
            }
            column(ShiptoContact; "Ship-to Contact")
            {
            }
            column(ShiptoCountryRegionCode; "Ship-to Country/Region Code")
            {
            }
            column(ShiptoCounty; "Ship-to County")
            {
            }
            column(ShiptoName; "Ship-to Name")
            {
            }
            column(ShiptoName2; "Ship-to Name 2")
            {
            }
            column(ShiptoPostCode; "Ship-to Post Code")
            {
            }
            column(ShortcutDimension1Code_Hdr; "Shortcut Dimension 1 Code")
            {
            }
            column(ShortcutDimension2Code_Hdr; "Shortcut Dimension 2 Code")
            {
            }
            column(ShortcutDimension4Code_Hdr; "Shortcut Dimension 4 Code")
            {
            }
            column(Status; Status)
            {
            }

            column(TaxAreaCode_Header; "Tax Area Code")
            {
            }
            column(TaxLiable_Header; "Tax Liable")
            {
            }
            column(TransactionModeCode; "Transaction Mode Code")
            {
            }
            column(USPaymentTerms; "US Payment Terms")
            {
            }
            column(USSalesOrderNo; "US Sales Order No.")
            {
            }
            column(VATAddressTelephone; "VAT Address & Telephone")
            {
            }
            column(VATBankNameAccount; "VAT Bank Name & Account")
            {
            }
            column(VATBaseDiscount; "VAT Base Discount %")
            {
            }
            column(VATBusPostingGroup_Header; "VAT Bus. Posting Group")
            {
            }
            column(VATContactInformation; "VAT Contact Information")
            {
            }
            column(VATCountryRegionCode; "VAT Country/Region Code")
            {
            }
            column(VATCustomerName; "VAT Customer Name")
            {
            }
            column(VATInvoiceMailAddress; "VAT Invoice Mail Address")
            {
            }
            column(VATRegistrationNo; "VAT Registration No.")
            {
            }
            column(WorkDescription; "Work Description")
            {
            }
            column(Document_Date; "Document Date")
            {
            }
            column(Salesperson_Code; "Salesperson Code")
            {
            }
            dataitem(Sales_Line; "Sales Line")
            {
                DataItemLink = "Document No." = Sales_Header."No.";
                DataItemTableFilter = "Document Type" = const(Order);
                column(Document_No; "Document No.")
                {
                }
                column(Type; Type)
                {

                }
                column(AllowInvoiceDisc; "Allow Invoice Disc.")
                {
                }
                column(AllowItemChargeAssignment; "Allow Item Charge Assignment")
                {
                }
                column(AllowLineDisc; "Allow Line Disc.")
                {
                }
                column(Amount; Amount)
                {
                }
                column(AmountIncludingVAT; "Amount Including VAT")
                {
                }
                column(ApplfromItemEntry; "Appl.-from Item Entry")
                {
                }
                column(AppltoItemEntry; "Appl.-to Item Entry")
                {
                }
                column("Area"; "Area")
                {
                }
                column(ATOWhseOutstandingQty; "ATO Whse. Outstanding Qty.")
                {
                }
                column(ATOWhseOutstdQtyBase; "ATO Whse. Outstd. Qty. (Base)")
                {
                }
                column(AttachedDocCount; "Attached Doc Count")
                {
                }
                column(AttachedtoLineNo; "Attached to Line No.")
                {
                }
                column(BilltoCustomerNo; "Bill-to Customer No.")
                {
                }
                column(BinCode; "Bin Code")
                {
                }
                column(BlanketOrderLineNo; "Blanket Order Line No.")
                {
                }
                column(BlanketOrderNo; "Blanket Order No.")
                {
                }
                column(BOMItemNo; "BOM Item No.")
                {
                }
                column(CheckedbySalesForce; "Checked by SalesForce")
                {
                }
                column(COC; COC)
                {
                }
                column(CompletelyShipped; "Completely Shipped")
                {
                }
                column(CopiedFromPostedDoc; "Copied From Posted Doc.")
                {
                }
                column(CurrencyCode; "Currency Code")
                {
                }
                column(CustomerDiscGroup; "Customer Disc. Group")
                {
                }
                column(CustomerPriceGroup; "Customer Price Group")
                {
                }
                column(DeferralCode; "Deferral Code")
                {
                }
                column(DepruntilFAPostingDate; "Depr. until FA Posting Date")
                {
                }
                column(DepreciationBookCode; "Depreciation Book Code")
                {
                }
                column(Description; Description)
                {
                }
                column(Description2; "Description 2")
                {
                }
                column(DimensionSetID; "Dimension Set ID")
                {
                }
                column(DocumentType; "Document Type")
                {
                }
                column(DropShipment; "Drop Shipment")
                {
                }
                column(DuplicateinDepreciationBook; "Duplicate in Depreciation Book")
                {
                }
                column(ESPofTotalAmount; "ESP % of Total Amount")
                {
                }
                column(ESPLCY; "ESP (LCY)")
                {
                }
                column(ESPDiscAmount; "ESP Disc.Amount")
                {
                }
                column(ESPLineAmount; "ESP Line Amount")
                {
                }
                column(ESPLineAmountLCY; "ESP Line Amount (LCY)")
                {
                }
                column(ESPNettLineAmount; "ESP Nett Line Amount")
                {
                }
                column(ExitPoint; "Exit Point")
                {
                }
                column(ExternalID; ExternalID)
                {
                }
                column(FAPostingDate; "FA Posting Date")
                {
                }
                column(GenBusPostingGroup; "Gen. Bus. Posting Group")
                {
                }
                column(GenProdPostingGroup; "Gen. Prod. Posting Group")
                {
                }
                column(GrossWeight; "Gross Weight")
                {
                }
                column(ICItemReferenceNo; "IC Item Reference No.")
                {
                }
                column(ICPartnerCode; "IC Partner Code")
                {
                }
                column(ICPartnerRefType; "IC Partner Ref. Type")
                {
                }
                column(ICPartnerReference; "IC Partner Reference")
                {
                }
                column(InvDiscAmounttoInvoice; "Inv. Disc. Amount to Invoice")
                {
                }
                column(InvDiscountAmount; "Inv. Discount Amount")
                {
                }
                column(ItemCategoryCode; "Item Category Code")
                {
                }
                column(ItemReferenceNo; "Item Reference No.")
                {
                }
                column(ItemReferenceType; "Item Reference Type")
                {
                }
                column(ItemReferenceTypeNo; "Item Reference Type No.")
                {
                }
                column(ItemReferenceUnitofMeasure; "Item Reference Unit of Measure")
                {
                }
                column(JobContractEntryNo; "Job Contract Entry No.")
                {
                }
                column(JobNo; "Job No.")
                {
                }
                column(JobTaskNo; "Job Task No.")
                {
                }
                column(LineAmount; "Line Amount")
                {
                }
                column(LineDiscount; "Line Discount %")
                {
                }
                column(LineDiscountAmount; "Line Discount Amount")
                {
                }
                column(LineDiscountCalculation; "Line Discount Calculation")
                {
                }
                column(LineNo; "Line No.")
                {
                }
                column(LocationCode; "Location Code")
                {
                }
                column(NetWeight; "Net Weight")
                {
                }
                column(NoofPeriods; "No. of Periods")
                {
                }
                column(Nonstock; Nonstock)
                {
                }
                column(OriginallyOrderedNo; "Originally Ordered No.")
                {
                }
                column(OriginallyOrderedVarCode; "Originally Ordered Var. Code")
                {
                }
                column(OutboundWhseHandlingTime; "Outbound Whse. Handling Time")
                {
                }
                column(OutofStockSubstitution; "Out-of-Stock Substitution")
                {
                }
                column(OutstandingAmount; "Outstanding Amount")
                {
                }
                column(OutstandingAmountLCY; "Outstanding Amount (LCY)")
                {
                }
                column(OutstandingQtyBase; "Outstanding Qty. (Base)")
                {
                }
                column(OutstandingQuantity; "Outstanding Quantity")
                {
                }
                column(Planned; Planned)
                {
                }
                column(PlannedDeliveryDate; "Planned Delivery Date")
                {
                }
                column(PlannedShipmentDate; "Planned Shipment Date")
                {
                }
                column(PmtDiscountAmount; "Pmt. Discount Amount")
                {
                }
                column(PostingDate; "Posting Date")
                {
                }
                column(PostingGroup; "Posting Group")
                {
                }
                column(Prepayment; "Prepayment %")
                {
                }
                column(PrepaymentAmount; "Prepayment Amount")
                {
                }
                column(PrepaymentLine; "Prepayment Line")
                {
                }
                column(PrepaymentTaxAreaCode; "Prepayment Tax Area Code")
                {
                }
                column(PrepaymentTaxGroupCode; "Prepayment Tax Group Code")
                {
                }
                column(PrepaymentTaxLiable; "Prepayment Tax Liable")
                {
                }
                column(PrepaymentVAT; "Prepayment VAT %")
                {
                }
                column(PrepaymentVATDifference; "Prepayment VAT Difference")
                {
                }
                column(PrepaymentVATIdentifier; "Prepayment VAT Identifier")
                {
                }
                column(PrepmtAmtDeducted; "Prepmt Amt Deducted")
                {
                }
                column(PrepmtAmttoDeduct; "Prepmt Amt to Deduct")
                {
                }
                column(PrepmtVATDiffDeducted; "Prepmt VAT Diff. Deducted")
                {
                }
                column(PrepmtVATDifftoDeduct; "Prepmt VAT Diff. to Deduct")
                {
                }
                column(PrepmtAmountInvLCY; "Prepmt. Amount Inv. (LCY)")
                {
                }
                column(PrepmtAmountInvInclVAT; "Prepmt. Amount Inv. Incl. VAT")
                {
                }
                column(PrepmtAmtInclVAT; "Prepmt. Amt. Incl. VAT")
                {
                }
                column(PrepmtAmtInv; "Prepmt. Amt. Inv.")
                {
                }
                column(PrepmtLineAmount; "Prepmt. Line Amount")
                {
                }
                column(PrepmtPmtDiscountAmount; "Prepmt. Pmt. Discount Amount")
                {
                }
                column(PrepmtVATAmountInvLCY; "Prepmt. VAT Amount Inv. (LCY)")
                {
                }
                column(PrepmtVATBaseAmt; "Prepmt. VAT Base Amt.")
                {
                }
                column(PrepmtVATCalcType; "Prepmt. VAT Calc. Type")
                {
                }
                column(PriceCalculationMethod; "Price Calculation Method")
                {
                }
                column(Pricedescription; "Price description")
                {
                }
                column(Profit; "Profit %")
                {
                }
                column(PromisedDeliveryDate; "Promised Delivery Date")
                {
                }
                column(PurchOrderLineNo; "Purch. Order Line No.")
                {
                }
                column(PurchaseOrderNo; "Purchase Order No.")
                {
                }
                column(PurchasingCode; "Purchasing Code")
                {
                }
                column(QtyAssigned; "Qty. Assigned")
                {
                }
                column(QtyInvoicedBase; "Qty. Invoiced (Base)")
                {
                }
                column(QtyperUnitofMeasure; "Qty. per Unit of Measure")
                {
                }
                column(QtyRoundingPrecision; "Qty. Rounding Precision")
                {
                }
                column(QtyRoundingPrecisionBase; "Qty. Rounding Precision (Base)")
                {
                }
                column(QtyShippedBase; "Qty. Shipped (Base)")
                {
                }
                column(QtyShippedNotInvdBase; "Qty. Shipped Not Invd. (Base)")
                {
                }
                column(QtyShippedNotInvoiced; "Qty. Shipped Not Invoiced")
                {
                }
                column(QtytoAsmtoOrderBase; "Qty. to Asm. to Order (Base)")
                {
                }
                column(QtytoAssembletoOrder; "Qty. to Assemble to Order")
                {
                }
                column(QtytoAssign; "Qty. to Assign")
                {
                }
                column(QtytoInvoice; "Qty. to Invoice")
                {
                }
                column(QtytoInvoiceBase; "Qty. to Invoice (Base)")
                {
                }
                column(QtytoShip; "Qty. to Ship")
                {
                }
                column(QtytoShipBase; "Qty. to Ship (Base)")
                {
                }
                column(Quantity; Quantity)
                {
                }
                column(QuantityBase; "Quantity (Base)")
                {
                }
                column(QuantityInvoiced; "Quantity Invoiced")
                {
                }
                column(QuantityShipped; "Quantity Shipped")
                {
                }
                column(RecalculateInvoiceDisc; "Recalculate Invoice Disc.")
                {
                }
                column(RequestedDeliveryDate; "Requested Delivery Date")
                {
                }
                column(Reserve; Reserve)
                {
                }
                column(ReservedQtyBase; "Reserved Qty. (Base)")
                {
                }
                column(ReservedQuantity; "Reserved Quantity")
                {
                }
                column(ResponsibilityCenter; "Responsibility Center")
                {
                }
                column(RetQtyRcdNotInvdBase; "Ret. Qty. Rcd. Not Invd.(Base)")
                {
                }
                column(ReturnQtyRcdNotInvd; "Return Qty. Rcd. Not Invd.")
                {
                }
                column(ReturnQtyReceived; "Return Qty. Received")
                {
                }
                column(ReturnQtyReceivedBase; "Return Qty. Received (Base)")
                {
                }
                column(ReturnQtytoReceive; "Return Qty. to Receive")
                {
                }
                column(ReturnQtytoReceiveBase; "Return Qty. to Receive (Base)")
                {
                }
                column(ReturnRcdNotInvd; "Return Rcd. Not Invd.")
                {
                }
                column(ReturnRcdNotInvdLCY; "Return Rcd. Not Invd. (LCY)")
                {
                }
                column(ReturnReasonCode; "Return Reason Code")
                {
                }
                column(ReturnReceiptLineNo; "Return Receipt Line No.")
                {
                }
                column(ReturnReceiptNo; "Return Receipt No.")
                {
                }
                column(ReturnsDeferralStartDate; "Returns Deferral Start Date")
                {
                }
                column(RevenueType; "Revenue Type")
                {
                }
                column(RRPofTotalAmount; "RRP % of Total Amount")
                {
                }
                column(RRPLCY; "RRP (LCY)")
                {
                }
                column(RRPDiscAmount; "RRP Disc.Amount")
                {
                }
                column(RRPLineAmount; "RRP Line Amount")
                {
                }
                column(RRPLineAmountLCY; "RRP Line Amount (LCY)")
                {
                }
                column(RRPNettLineAmount; "RRP Nett Line Amount")
                {
                }
                column(SalesOrderLineNo; "Sales Order Line No.")
                {
                }
                column(SalesOrderNo; "Sales Order No.")
                {
                }
                column(SelltoCustomerNo; "Sell-to Customer No.")
                {
                }
                column(ShipmentDate; "Shipment Date")
                {
                }
                column(ShipmentLineNo; "Shipment Line No.")
                {
                }
                column(ShipmentNo; "Shipment No.")
                {
                }
                column(ShippedNotInvLCYNoVAT; "Shipped Not Inv. (LCY) No VAT")
                {
                }
                column(ShippedNotInvoiced; "Shipped Not Invoiced")
                {
                }
                column(ShippedNotInvoicedLCY; "Shipped Not Invoiced (LCY)")
                {
                }
                column(ShippingAgentCode; "Shipping Agent Code")
                {
                }
                column(ShippingAgentServiceCode; "Shipping Agent Service Code")
                {
                }
                column(ShippingTime; "Shipping Time")
                {
                }
                column(ShortcutDimension1Code; "Shortcut Dimension 1 Code")
                {
                }
                column(ShortcutDimension2Code; "Shortcut Dimension 2 Code")
                {
                }
                column(ShortcutDimension4Code; "Shortcut Dimension 4 Code")
                {
                }
                column(SortingNo; "Sorting No.")
                {
                }
                column(SpecialOrder; "Special Order")
                {
                }
                column(SpecialOrderPurchLineNo; "Special Order Purch. Line No.")
                {
                }
                column(SpecialOrderPurchaseNo; "Special Order Purchase No.")
                {
                }
                column(StartingDate; "Starting Date")
                {
                }
                column(SubstitutionAvailable; "Substitution Available")
                {
                }
                column(Subtype; Subtype)
                {
                }
                column(SystemCreatedEntry; "System-Created Entry")
                {
                }
                column(SystemCreatedAt; SystemCreatedAt)
                {
                }
                column(SystemCreatedBy; SystemCreatedBy)
                {
                }
                column(SystemId; SystemId)
                {
                }
                column(SystemModifiedAt; SystemModifiedAt)
                {
                }
                column(SystemModifiedBy; SystemModifiedBy)
                {
                }
                column(TaxAreaCode; "Tax Area Code")
                {
                }
                column(TaxCategory; "Tax Category")
                {
                }
                column(TaxGroupCode; "Tax Group Code")
                {
                }
                column(TaxLiable; "Tax Liable")
                {
                }
                column(TransactionSpecification; "Transaction Specification")
                {
                }
                column("TransactionType"; "Transaction Type")
                {
                }
                column(TransportMethod; "Transport Method")
                {
                }
                column(UnitCost; "Unit Cost")
                {
                }
                column(UnitCostLCY; "Unit Cost (LCY)")
                {
                }
                column(UnitofMeasure; "Unit of Measure")
                {
                }
                column(UnitofMeasureCode; "Unit of Measure Code")
                {
                }
                column(UnitPrice; "Unit Price")
                {
                }
                column(UnitVolume; "Unit Volume")
                {
                }
                column(UnitsperParcel; "Units per Parcel")
                {
                }
                column(UseDuplicationList; "Use Duplication List")
                {
                }
                column(VariantCode; "Variant Code")
                {
                }
                column(VAT; "VAT %")
                {
                }
                column(VATBaseAmount; "VAT Base Amount")
                {
                }
                column(VATBusPostingGroup; "VAT Bus. Posting Group")
                {
                }
                column(VATCalculationType; "VAT Calculation Type")
                {
                }
                column(VATClauseCode; "VAT Clause Code")
                {
                }
                column(VATDifference; "VAT Difference")
                {
                }
                column(VATIdentifier; "VAT Identifier")
                {
                }
                column(VATProdPostingGroup; "VAT Prod. Posting Group")
                {
                }
                column(WhseOutstandingQty; "Whse. Outstanding Qty.")
                {
                }
                column(WhseOutstandingQtyBase; "Whse. Outstanding Qty. (Base)")
                {
                }
                column(WorkTypeCode; "Work Type Code")
                {
                }
            }
        }
    }
}

