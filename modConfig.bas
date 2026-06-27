Attribute VB_Name = "modConfig"
Option Explicit

' =====================================================
' ===== CONFIGURATION CONSTANTS =====
' =====================================================

' ===== APP VERSION =====
Public Const APP_VERSION As String = "1.5.0"
Public Const APP_BUILD   As String = "2026-06-27"

' ===== ROW CONFIGURATION =====
Public Const START_ROW As Integer = 14
Public Const END_ROW As Integer = 31

' ===== SUMMARY ROW CONFIGURATION =====
Public Const ROW_DISCOUNT_LABEL As Integer = 32
Public Const ROW_DISCOUNT_AMOUNT As Integer = 33
Public Const ROW_SUBTOTAL As Integer = 34
Public Const ROW_VAT As Integer = 35
Public Const ROW_TOTAL As Integer = 36

' ===== VAT RATE =====
Public Const VAT_RATE As Double = 0.15

' ===== QUOTE / INVOICE CELL ADDRESSES =====
' Centralised so layout changes only need one edit.
Public Const CELL_DOC_NUMBER     As String = "G8"
Public Const CELL_DOC_DATE       As String = "G7"
Public Const CELL_CUSTOMER       As String = "C7"
Public Const CELL_PATIENT        As String = "G12"
Public Const CELL_APPLIANCE      As String = "D12"
Public Const CELL_SUMMARY_SUBTOT As String = "H34"
Public Const CELL_SUMMARY_DISC   As String = "H33"
Public Const CELL_SUMMARY_VAT    As String = "H35"
Public Const CELL_SUMMARY_TOTAL  As String = "H36"
Public Const CELL_DISCOUNT_PCT   As String = "C32"

' ===== SHEET NAMES (Main Workbook) =====
Public Const SHEET_MENU As String = "Menu"
Public Const SHEET_QUOTE As String = "Quote"
Public Const SHEET_INVOICE As String = "Invoice"
Public Const SHEET_PRICELIST As String = "PriceList"
Public Const SHEET_SETTINGS As String = "Settings"
Public Const SHEET_CUSTOMERS As String = "Customers"
Public Const SHEET_STATEMENT As String = "Statement"

' ===== TRACKING SHEET NAMES =====
Public Const TRACK_SHEET_QUOTES As String = "Quotes"
Public Const TRACK_SHEET_INVOICES As String = "Invoices"

' ===== DOCUMENT NUMBER PREFIXES =====
Public Const QUOTE_PREFIX   As String = "WA"
Public Const INVOICE_PREFIX As String = "INV"

' ===== STATUS VALUES - QUOTES =====
Public Const STATUS_QUOTE_OPEN As String = "Open"
Public Const STATUS_QUOTE_CONVERTED As String = "Converted"
Public Const STATUS_QUOTE_CANCELLED As String = "Cancelled"
Public Const STATUS_QUOTE_EXPIRED As String = "Expired"

' ===== STATUS VALUES - INVOICES =====
Public Const STATUS_INVOICE_UNPAID As String = "Unpaid"
Public Const STATUS_INVOICE_PARTIAL As String = "Partial"
Public Const STATUS_INVOICE_PAID As String = "Paid"
Public Const STATUS_INVOICE_OVERDUE As String = "Overdue"
Public Const STATUS_INVOICE_CANCELLED As String = "Cancelled"

' =====================================================
' ===== QUOTE TRACKER COLUMNS =====
' =====================================================
Public Const COL_Q_NUMBER       As Long = 1
Public Const COL_Q_DATE         As Long = 2
Public Const COL_Q_CUSTOMER     As Long = 3
Public Const COL_Q_PATIENT      As Long = 4
Public Const COL_Q_APPLIANCE    As Long = 5
Public Const COL_Q_SUBTOTAL     As Long = 6
Public Const COL_Q_DISCOUNT     As Long = 7
Public Const COL_Q_VAT          As Long = 8
Public Const COL_Q_TOTAL        As Long = 9
Public Const COL_Q_STATUS       As Long = 10
Public Const COL_Q_INVOICENUM   As Long = 11
Public Const COL_Q_FILEPATH     As Long = 12
Public Const COL_Q_LASTMOD      As Long = 13

' =====================================================
' ===== INVOICE TRACKER COLUMNS =====
' =====================================================
Public Const COL_I_NUMBER       As Long = 1
Public Const COL_I_QUOTENUM     As Long = 2
Public Const COL_I_DATE         As Long = 3
Public Const COL_I_CUSTOMER     As Long = 4
Public Const COL_I_PATIENT      As Long = 5
Public Const COL_I_APPLIANCE    As Long = 6
Public Const COL_I_SUBTOTAL     As Long = 7
Public Const COL_I_DISCOUNT     As Long = 8
Public Const COL_I_VAT          As Long = 9
Public Const COL_I_TOTAL        As Long = 10
Public Const COL_I_STATUS       As Long = 11
Public Const COL_I_PAIDDATE     As Long = 12
Public Const COL_I_FILEPATH     As Long = 13
Public Const COL_I_LASTMOD      As Long = 14
Public Const COL_I_CREDITAMT    As Long = 15
Public Const COL_I_CREDITREASON As Long = 16
Public Const COL_I_PAIDAMOUNT   As Long = 17

' =====================================================
' ===== MENU SHEET LAYOUT =====
' =====================================================
' Open Quotes panel rows on Menu sheet
Public Const MENU_OPEN_QUOTES_HEADER_ROW As Long = 29
Public Const MENU_OPEN_QUOTES_FIRST_DATA As Long = 31
Public Const MENU_OPEN_QUOTES_MAX_ROWS   As Long = 15

' Derived constants so summary/tip rows can't collide with the Unpaid panel
Public Const MENU_OPEN_QUOTES_LAST_DATA  As Long = MENU_OPEN_QUOTES_FIRST_DATA + MENU_OPEN_QUOTES_MAX_ROWS - 1
Public Const MENU_OPEN_QUOTES_TOTAL_ROW  As Long = MENU_OPEN_QUOTES_LAST_DATA + 1
Public Const MENU_OPEN_QUOTES_TIP_ROW    As Long = MENU_OPEN_QUOTES_TOTAL_ROW + 1

' --- Open Quotes Customer filter cell on Menu sheet ---
Public Const MENU_OPEN_QUOTES_CUST_CELL As String = "C27"

' --- NEW: Open Quotes pagination UI cells (live in row 28, just above the table header) ---
'   Prev / Next are user-clickable cells (double-click in Sheet10 code-behind).
'   Page label sits between them so users see "Page 2 of 5" at a glance.
'   Current page is stored in a hidden helper cell on the Menu sheet so it
'   survives between refreshes WITHOUT polluting module-level state.
Public Const MENU_OPEN_QUOTES_PREV_CELL  As String = "F28"   ' double-click = previous page
Public Const MENU_OPEN_QUOTES_PAGE_CELL  As String = "G28"   ' shows "Page X of Y"
Public Const MENU_OPEN_QUOTES_NEXT_CELL  As String = "H28"   ' double-click = next page
Public Const MENU_OPEN_QUOTES_STATE_CELL As String = "K28"   ' hidden — stores current page number (1-based)

' Working sheets that the menu shows/hides
Public Const MENU_WORKING_SHEETS As String = "Quote|Invoice|Statement"
' Always-hidden sheets while menu is the dashboard
Public Const MENU_BACKGROUND_SHEETS As String = "PriceList|Customers|Settings"

' =====================================================
' ===== UNPAID INVOICES PANEL LAYOUT =====
' Derived from Open Quotes layout so the two panels never overlap.
' =====================================================
Public Const MENU_UNPAID_INV_TITLE_ROW  As Long = MENU_OPEN_QUOTES_FIRST_DATA + MENU_OPEN_QUOTES_MAX_ROWS + 3   ' 49
Public Const MENU_UNPAID_INV_FILTER_ROW As Long = MENU_UNPAID_INV_TITLE_ROW + 1                                  ' 50
Public Const MENU_UNPAID_INV_CUST_ROW   As Long = MENU_UNPAID_INV_TITLE_ROW + 2                                  ' 51
Public Const MENU_UNPAID_INV_HEADER_ROW As Long = MENU_UNPAID_INV_TITLE_ROW + 3                                  ' 52
Public Const MENU_UNPAID_INV_FIRST_DATA As Long = MENU_UNPAID_INV_TITLE_ROW + 5                                  ' 54
Public Const MENU_UNPAID_INV_MAX_ROWS   As Long = 15

' --- Unpaid Invoices filter cells (string addresses depend on the rows above) ---
Public Const MENU_UNPAID_INV_FROM_CELL      As String = "C50"
Public Const MENU_UNPAID_INV_TO_CELL        As String = "F50"
Public Const MENU_UNPAID_INV_FROM_PICK_CELL As String = "D50"
Public Const MENU_UNPAID_INV_TO_PICK_CELL   As String = "G50"
Public Const MENU_UNPAID_INV_APPLY_CELL     As String = "H50"
Public Const MENU_UNPAID_INV_CUST_CELL      As String = "C51"

' --- NEW: Unpaid Invoices pagination UI cells (live in row 51, on the header row, far right) ---
'   Same pattern as Open Quotes pagination.
Public Const MENU_UNPAID_INV_PREV_CELL  As String = "F51"
Public Const MENU_UNPAID_INV_PAGE_CELL  As String = "G51"
Public Const MENU_UNPAID_INV_NEXT_CELL  As String = "H51"
Public Const MENU_UNPAID_INV_STATE_CELL As String = "K51"
