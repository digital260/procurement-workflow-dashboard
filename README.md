# 📊 Procurement Workflow Dashboard

A comprehensive Google Sheets-based procurement workflow automation system that streamlines the entire procurement process from request to payment.

## ✨ Features

- **📋 Procurement Requests**: Submit and track procurement requests
- **📦 Material Tracking**: Monitor shipments in real-time
- **✅ Quality Checks**: Record material inspection and quality decisions
- **💳 Payment Processing**: Process vendor payments
- **📊 Live Dashboard**: Real-time overview of all procurement metrics
- **🔔 Email Notifications**: Automatic email updates for requestors
- **🏷️ Auto-Generated IDs**: Unique identifiers for all transactions
- **📈 Audit Trail**: Complete history of all procurement activities

## 📋 Workflow Stages

1. **START** → Yellow button
2. **PROCUREMENT BUDGET** → Check if item in budget
3. **CHECK EXISTING PURCHASE** → If YES/NO decision
4. **IDENTIFY VENDOR** → Select and get quotes
5. **MANAGER APPROVAL** → Approval decision
6. **MATERIAL TRANSIT** → Track package with courier
7. **MATERIAL RECEIVED** → Confirm receipt
8. **QUALITY CHECK** → PASS/FAIL inspection
   - If PASS → Proceed to GRNN
   - If FAIL → Raise Debit Note
9. **GENERATE GRNN** → Goods Receipt Note
10. **STOCK ENTRY** → Add to inventory
11. **3-WAY MATCH** → Match PO vs GRN vs Invoice
12. **BILL AGAINST PO** → Process invoice
13. **PROCESS PAYMENT** → Payment decision
14. **GENERATE PAYMENT** → Create payment
15. **END** → Complete

## 📊 Sheet Structure

### 1. **Dashboard**
- Live metrics and KPIs
- Total requests, approvals, transit status
- Payment tracking
- Amount summaries

### 2. **Requests**
- Request ID
- Date & Requestor
- Amount & Description
- Status (Pending/Approved)
- Manager approval tracking

### 3. **PO Master**
- Purchase Order details
- Vendor information
- PO amounts and delivery dates
- Status tracking

### 4. **Material Transit**
- Transit IDs & Tracking numbers
- Dispatch and delivery dates
- Courier and location info
- Real-time status updates

### 5. **Quality Check**
- Quality inspection records
- PASS/FAIL decisions
- Defect documentation
- GRNN generation status

### 6. **Payment**
- Payment IDs & Invoice numbers
- Vendor payment tracking
- Payment methods
- Approval records

### 7. **Vendor Master**
- Vendor details
- Contact information
- Bank details
- Payment terms

## 🚀 Setup Instructions

### Prerequisites
- Google Account
- Access to Google Sheets
- Editor access to create AppScripts

### Installation

1. **Create a new Google Sheet**
   - Go to [Google Sheets](https://sheets.google.com)
   - Click "Create new spreadsheet"

2. **Add the AppScript Code**
   - In your Sheet, go to `Extensions` → `Apps Script`
   - Replace the default code with `Code.gs` from this repository
   - Save the file

3. **Deploy as Web App**
