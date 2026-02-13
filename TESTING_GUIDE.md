# 🧪 System Testing Guide

Since you've effectively created an invoice, follow this sequence to test the rest of the system.

## 1. Test Payment Recording (Partial)
*Goal: Record a partial payment and verify balance update.*

1.  Click **RECORD PAYMENT** on the Dashboard.
2.  **Invoice No**: Enter the invoice number you just created (e.g., `INV-2026-0001`).
3.  **Date**: Leave as default (today).
4.  **Amount**: Enter **half** the total amount (e.g., if total is 10,000, enter 5,000).
5.  **Payment Method**: Select "M-Pesa" or "Bank Transfer".
6.  **Reference**: Enter a test ref (e.g., `REF123`).
7.  **Click OK**.
    *   ✅ **Check**: Go to **Transactions** sheet. The "Amount Paid" column should increase, and "Balance" should decrease. Status should be "Partial".
    *   ✅ **Check**: Go to **PaymentLog** sheet. A new row should appear.

## 2. Test Receipt Generation
*Goal: Generate a receipt for the payment.*

1.  Click **GENERATE RECEIPT**.
2.  **Invoice No**: Enter `INV-2026-0001`.
3.  **Click OK**.
    *   ✅ **Check**: You should be taken to the **Receipt_Template** sheet.
    *   ✅ **Check**: It should show "Amount Paid" and the remaining "Balance".
    *   ✅ **Check**: A PDF should be saved in your folder (under `Receipts/`).

## 3. Test Full Payment
*Goal: Clear the remaining balance.*

1.  Click **RECORD PAYMENT** again.
2.  **Invoice No**: `INV-2026-0001`.
3.  **Amount**: Enter the **remaining balance**.
4.  **Reference**: `REF124`.
5.  **Click OK**.
    *   ✅ **Check**: Message should say "Invoice Fully Paid".
    *   ✅ **Check**: In **Transactions** sheet, Status should change to **"Paid"**.

## 4. Test ETR Generation
*Goal: Generate a fiscal receipt.*

1.  Click **ETR RECEIPT** on the Dashboard.
2.  **Invoice No**: `INV-2026-0001`.
3.  **Click OK**.
    *   ✅ **Check**: You should be taken to **ETR_Template**.
    *   ✅ **Check**: The layout should look like a long, narrow thermal receipt.
    *   ✅ **Check**: A PDF should be saved in `ETR/`.

## 5. Test Reports & Dashboard
*Goal: Verify stats update automatically.*

1.  Click **REFRESH** (or View -> Refresh Dashboard if you added that button, otherwise just look at the cards).
2.  **Total Revenue**: Should include the logic for paid/unpaid.
3.  **Tax Collected**: Should reflect the tax from your invoice.
4.  **Outstanding**: Should be **0.00** (since you fully paid it).
5.  **Recent Activity**: Should list your invoice and its status.
6.  Go to **TAX SUMMARY** (via button):
    *   Check if the numbers match your invoice tax.

## 6. Test PDF Export (Optional)
*Goal: Manually export a specific sheet.*

1.  Go to the **Invoice_Template** sheet (navigate manually or click New Invoice -> Cancel -> verify sheet).
2.  Click **EXPORT PDF** on the Dashboard (or run macro `ExportActiveSheetToPDF`).
3.  Check the `PDF_Exports` folder.

## 7. Test Settings Sheet (Configuration)
*Goal: Verify that changing settings updates the system instantly.*

### Example A: Change Company Name
1.  Go to **Settings** sheet.
2.  Change **Cell B2** (Company Name) to something visible like "TEST CORP LTD".
3.  Go to **Invoice_Template** sheet.
4.  ✅ **Check**: The header at Cell E9 (or wherever the company name appears) should now say "TEST CORP LTD".

### Example B: Add Payment Method
1.  Go to **Settings** sheet.
2.  Scroll down to **Row 32** (Payment Methods).
3.  Add a new method in the list (e.g., "Crypto").
4.  Go to **Dashboard** -> **RECORD PAYMENT**.
5.  ✅ **Check**: Open the "Method" dropdown. "Crypto" should now be an option.
