/**
 * PROCUREMENT WORKFLOW DASHBOARD
 * Complete implementation for Google Sheets
 * Author: Procurement Team
 * Version: 1.0
 */

// ============================================================================
// INITIALIZATION & MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Procurement Workflow')
    .addItem('📋 Start New Procurement', 'openProcurementForm')
    .addItem('📦 Track Material', 'openTrackingForm')
    .addItem('✅ Quality Check', 'openQualityForm')
    .addItem('💳 Process Payment', 'openPaymentForm')
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🔧 Setup Sheets', 'initializeSheets')
    .addToUi();
}

// ============================================================================
// INITIALIZATION FUNCTIONS
// ============================================================================

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Dashboard', 'Requests', 'PO Master', 'Material Transit', 'Quality Check', 'Payment', 'Vendor Master'];
  
  sheetNames.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  
  // Setup Dashboard
  setupDashboard();
  
  // Setup Requests sheet
  setupRequestsSheet();
  
  // Setup PO Master sheet
  setupPOMasterSheet();
  
  // Setup Material Transit sheet
  setupMaterialTransitSheet();
  
  // Setup Quality Check sheet
  setupQualityCheckSheet();
  
  // Setup Payment sheet
  setupPaymentSheet();
  
  // Setup Vendor Master sheet
  setupVendorMasterSheet();
  
  SpreadsheetApp.getUi().alert('✅ All sheets initialized successfully!');
}

function setupDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  sheet.clear();
  
  const headers = [
    'PROCUREMENT WORKFLOW DASHBOARD',
    '',
    new Date()
  ];
  
  sheet.appendRow(headers);
  sheet.appendRow(['']);
  sheet.appendRow(['Metric', 'Count', 'Status']);
  sheet.appendRow(['Total Requests', '=COUNTA(Requests!A2:A)', '']);
  sheet.appendRow(['Pending Approval', '=COUNTIF(Requests!D2:D,"Pending")', '']);
  sheet.appendRow(['Approved POs', '=COUNTIF(Requests!D2:D,"Approved")', '']);
  sheet.appendRow(['In Transit', '=COUNTIF(\'Material Transit\'!E2:E,"In Transit")', '']);
  sheet.appendRow(['Quality Passed', '=COUNTIF(\'Quality Check\'!D2:D,"PASS")', '']);
  sheet.appendRow(['Payments Pending', '=COUNTIF(Payment!F2:F,"Pending")', '']);
  sheet.appendRow(['Total Amount', '=SUM(\'PO Master\'!F2:F)', '']);
  
  // Format headers
  const range = sheet.getRange(1, 1, 1, 3);
  range.setFontSize(14).setFontWeight('bold').setBackground('#1f4e78').setFontColor('white');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
}

function setupRequestsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  sheet.clear();
  
  const headers = [
    'Request ID',
    'Date',
    'Requestor',
    'Status',
    'Amount',
    'Description',
    'Manager Approval',
    'Approval Date',
    'Notes'
  ];
  
  sheet.appendRow(headers);
  
  // Format headers
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
}

function setupPOMasterSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PO Master');
  sheet.clear();
  
  const headers = [
    'PO Number',
    'Request ID',
    'Vendor Name',
    'Vendor Contact',
    'PO Date',
    'Amount',
    'Items',
    'Delivery Date',
    'Status',
    'Created By'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#2196F3').setFontColor('white');
  
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
}

function setupMaterialTransitSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Material Transit');
  sheet.clear();
  
  const headers = [
    'Transit ID',
    'PO Number',
    'Tracking Number',
    'Dispatch Date',
    'Status',
    'Expected Delivery',
    'Actual Delivery',
    'Courier',
    'Location',
    'Notes'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#FF9800').setFontColor('white');
}

function setupQualityCheckSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Quality Check');
  sheet.clear();
  
  const headers = [
    'Check ID',
    'PO Number',
    'Received Date',
    'Quality Status',
    'Defects Found',
    'Checked By',
    'Check Date',
    'Action Required',
    'GRNN Generated',
    'Notes'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#9C27B0').setFontColor('white');
}

function setupPaymentSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payment');
  sheet.clear();
  
  const headers = [
    'Payment ID',
    'PO Number',
    'Invoice Number',
    'Vendor Name',
    'Amount',
    'Status',
    'Payment Date',
    'Payment Method',
    'Reference',
    'Approval By'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#F44336').setFontColor('white');
}

function setupVendorMasterSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vendor Master');
  sheet.clear();
  
  const headers = [
    'Vendor ID',
    'Vendor Name',
    'Contact Person',
    'Email',
    'Phone',
    'Address',
    'City',
    'Bank Details',
    'Payment Terms',
    'Active'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#00BCD4').setFontColor('white');
}

// ============================================================================
// FORM DIALOGS
// ============================================================================

function openProcurementForm() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1f4e78; margin-bottom: 10px; font-size: 24px; }
        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; color: #333; }
        input[type="text"], input[type="email"], input[type="number"], input[type="date"], select, textarea {
          width: 100%;
          padding: 12px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 14px;
          font-family: Arial;
        }
        textarea { resize: vertical; min-height: 80px; }
        button { 
          width: 100%; 
          padding: 12px; 
          background: #4CAF50; 
          color: white; 
          border: none; 
          border-radius: 4px; 
          font-size: 16px; 
          font-weight: bold;
          cursor: pointer;
          margin-top: 20px;
        }
        button:hover { background: #45a049; }
        .info { background: #e3f2fd; padding: 12px; border-radius: 4px; margin-bottom: 20px; color: #1565c0; font-size: 13px; }
        .success { color: #4CAF50; display: none; margin-top: 10px; }
        .error { color: #f44336; display: none; margin-top: 10px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🛒 New Procurement Request</h1>
        <p class="subtitle">Start a new procurement workflow</p>
        
        <div class="info">
          <strong>ℹ️ Note:</strong> Fill in all fields to submit the request for manager approval.
        </div>
        
        <form id="procurementForm">
          <div class="form-group">
            <label for="requestor">Requestor Name *</label>
            <input type="text" id="requestor" placeholder="Your name" required>
          </div>
          
          <div class="form-group">
            <label for="email">Email *</label>
            <input type="email" id="email" placeholder="your.email@company.com" required>
          </div>
          
          <div class="form-group">
            <label for="amount">Amount (₹) *</label>
            <input type="number" id="amount" placeholder="50000" min="0" step="0.01" required>
          </div>
          
          <div class="form-group">
            <label for="description">Description *</label>
            <textarea id="description" placeholder="Describe what you need to procure" required></textarea>
          </div>
          
          <div class="form-group">
            <label for="vendor">Preferred Vendor (Optional)</label>
            <input type="text" id="vendor" placeholder="Vendor name">
          </div>
          
          <div class="form-group">
            <label for="deliveryDate">Required Delivery Date *</label>
            <input type="date" id="deliveryDate" required>
          </div>
          
          <div class="form-group">
            <label for="notes">Additional Notes</label>
            <textarea id="notes" placeholder="Any special requirements"></textarea>
          </div>
          
          <button type="submit">📤 Submit Request</button>
          <div class="success" id="successMsg">✅ Request submitted successfully!</div>
          <div class="error" id="errorMsg"></div>
        </form>
      </div>
      
      <script>
        document.getElementById('procurementForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const data = {
            requestor: document.getElementById('requestor').value,
            email: document.getElementById('email').value,
            amount: document.getElementById('amount').value,
            description: document.getElementById('description').value,
            vendor: document.getElementById('vendor').value,
            deliveryDate: document.getElementById('deliveryDate').value,
            notes: document.getElementById('notes').value
          };
          
          google.script.run.withSuccessHandler(function(result) {
            document.getElementById('successMsg').style.display = 'block';
            document.getElementById('procurementForm').reset();
            setTimeout(() => google.script.host.close(), 2000);
          }).withFailureHandler(function(error) {
            document.getElementById('errorMsg').textContent = 'Error: ' + error;
            document.getElementById('errorMsg').style.display = 'block';
          }).submitProcurementRequest(data);
        });
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '🛒 New Procurement Request');
}

function openTrackingForm() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: #FF9800; margin-bottom: 10px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], select { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        button { width: 100%; padding: 12px; background: #FF9800; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 20px; }
        button:hover { background: #e68900; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>📦 Track Material Transit</h1>
        
        <div class="form-group">
          <label for="poNumber">PO Number *</label>
          <input type="text" id="poNumber" placeholder="Enter PO number" required>
        </div>
        
        <div class="form-group">
          <label for="trackingNumber">Tracking Number *</label>
          <input type="text" id="trackingNumber" placeholder="Enter tracking number" required>
        </div>
        
        <div class="form-group">
          <label for="status">Status *</label>
          <select id="status" required>
            <option>Select Status</option>
            <option>In Transit</option>
            <option>Out for Delivery</option>
            <option>Delivered</option>
            <option>Delayed</option>
          </select>
        </div>
        
        <div class="form-group">
          <label for="location">Current Location</label>
          <input type="text" id="location" placeholder="Warehouse / City">
        </div>
        
        <div class="form-group">
          <label for="expectedDelivery">Expected Delivery Date</label>
          <input type="date" id="expectedDelivery">
        </div>
        
        <button type="button" onclick="submitTracking()">📤 Update Transit Status</button>
      </div>
      
      <script>
        function submitTracking() {
          const data = {
            poNumber: document.getElementById('poNumber').value,
            trackingNumber: document.getElementById('trackingNumber').value,
            status: document.getElementById('status').value,
            location: document.getElementById('location').value,
            expectedDelivery: document.getElementById('expectedDelivery').value
          };
          
          google.script.run.withSuccessHandler(function() {
            alert('✅ Transit tracking updated!');
            google.script.host.close();
          }).submitMaterialTransit(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '📦 Track Material');
}

function openQualityForm() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: #9C27B0; margin-bottom: 10px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], select, textarea { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        button { width: 100%; padding: 12px; background: #9C27B0; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 20px; }
        button:hover { background: #7b1fa2; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>✅ Quality Check</h1>
        
        <div class="form-group">
          <label for="poNumber">PO Number *</label>
          <input type="text" id="poNumber" placeholder="Enter PO number" required>
        </div>
        
        <div class="form-group">
          <label for="qualityStatus">Quality Status *</label>
          <select id="qualityStatus" required>
            <option>Select Status</option>
            <option>PASS</option>
            <option>FAIL</option>
            <option>PARTIAL</option>
          </select>
        </div>
        
        <div class="form-group">
          <label for="defects">Defects Found (if any)</label>
          <textarea id="defects" placeholder="Describe any defects"></textarea>
        </div>
        
        <div class="form-group">
          <label for="checkedBy">Checked By *</label>
          <input type="text" id="checkedBy" placeholder="Your name" required>
        </div>
        
        <div class="form-group">
          <label for="notes">Additional Notes</label>
          <textarea id="notes" placeholder="Any observations"></textarea>
        </div>
        
        <button type="button" onclick="submitQuality()">✅ Submit Quality Check</button>
      </div>
      
      <script>
        function submitQuality() {
          const data = {
            poNumber: document.getElementById('poNumber').value,
            qualityStatus: document.getElementById('qualityStatus').value,
            defects: document.getElementById('defects').value,
            checkedBy: document.getElementById('checkedBy').value,
            notes: document.getElementById('notes').value
          };
          
          google.script.run.withSuccessHandler(function() {
            alert('✅ Quality check recorded!');
            google.script.host.close();
          }).submitQualityCheck(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '✅ Quality Check');
}

function openPaymentForm() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: #F44336; margin-bottom: 10px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], input[type="number"], select, textarea { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        button { width: 100%; padding: 12px; background: #F44336; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 20px; }
        button:hover { background: #da190b; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>💳 Process Payment</h1>
        
        <div class="form-group">
          <label for="poNumber">PO Number *</label>
          <input type="text" id="poNumber" placeholder="Enter PO number" required>
        </div>
        
        <div class="form-group">
          <label for="invoiceNumber">Invoice Number *</label>
          <input type="text" id="invoiceNumber" placeholder="Enter invoice number" required>
        </div>
        
        <div class="form-group">
          <label for="vendorName">Vendor Name *</label>
          <input type="text" id="vendorName" placeholder="Vendor name" required>
        </div>
        
        <div class="form-group">
          <label for="amount">Amount (₹) *</label>
          <input type="number" id="amount" placeholder="Amount" min="0" step="0.01" required>
        </div>
        
        <div class="form-group">
          <label for="paymentMethod">Payment Method *</label>
          <select id="paymentMethod" required>
            <option>Select Method</option>
            <option>Bank Transfer</option>
            <option>Check</option>
            <option>Credit Card</option>
            <option>Cash</option>
          </select>
        </div>
        
        <div class="form-group">
          <label for="reference">Reference/Cheque Number</label>
          <input type="text" id="reference" placeholder="Reference number">
        </div>
        
        <div class="form-group">
          <label for="notes">Notes</label>
          <textarea id="notes" placeholder="Payment notes"></textarea>
        </div>
        
        <button type="button" onclick="submitPayment()">💳 Process Payment</button>
      </div>
      
      <script>
        function submitPayment() {
          const data = {
            poNumber: document.getElementById('poNumber').value,
            invoiceNumber: document.getElementById('invoiceNumber').value,
            vendorName: document.getElementById('vendorName').value,
            amount: document.getElementById('amount').value,
            paymentMethod: document.getElementById('paymentMethod').value,
            reference: document.getElementById('reference').value,
            notes: document.getElementById('notes').value
          };
          
          google.script.run.withSuccessHandler(function() {
            alert('✅ Payment processed!');
            google.script.host.close();
          }).submitPayment(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '💳 Process Payment');
}

// ============================================================================
// SUBMISSION HANDLERS
// ============================================================================

function submitProcurementRequest(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Requests');
  
  const requestId = 'REQ-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const row = [
    requestId,
    new Date(),
    data.requestor,
    'Pending',
    data.amount,
    data.description,
    '',
    '',
    data.notes
  ];
  
  sheet.appendRow(row);
  
  // Send notification email
  const mailTo = data.email;
  const subject = 'Procurement Request Submitted - ' + requestId;
  const message = `
Dear ${data.requestor},

Your procurement request has been submitted successfully.

Request Details:
- Request ID: ${requestId}
- Amount: ₹${data.amount}
- Description: ${data.description}
- Required Delivery: ${data.deliveryDate}
- Status: Pending Manager Approval

You will receive updates as the request progresses through the workflow.

Best regards,
Procurement Team
  `;
  
  GmailApp.sendEmail(mailTo, subject, message);
  
  Logger.log('Request submitted: ' + requestId);
}

function submitMaterialTransit(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Material Transit');
  
  const transitId = 'TRAN-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const row = [
    transitId,
    data.poNumber,
    data.trackingNumber,
    new Date(),
    data.status,
    data.expectedDelivery,
    '',
    '',
    data.location,
    ''
  ];
  
  sheet.appendRow(row);
  Logger.log('Transit tracking updated: ' + transitId);
}

function submitQualityCheck(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Quality Check');
  
  const checkId = 'QC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const row = [
    checkId,
    data.poNumber,
    new Date(),
    data.qualityStatus,
    data.defects,
    data.checkedBy,
    new Date(),
    data.qualityStatus === 'FAIL' ? 'Raise Debit Note' : 'Proceed to Payment',
    data.qualityStatus === 'PASS' ? 'YES' : 'PENDING',
    data.notes
  ];
  
  sheet.appendRow(row);
  Logger.log('Quality check recorded: ' + checkId);
}

function submitPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Payment');
  
  const paymentId = 'PAY-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const row = [
    paymentId,
    data.poNumber,
    data.invoiceNumber,
    data.vendorName,
    data.amount,
    'Processed',
    new Date(),
    data.paymentMethod,
    data.reference,
    Session.getActiveUser().getEmail()
  ];
  
  sheet.appendRow(row);
  Logger.log('Payment processed: ' + paymentId);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  
  // Recalculate all formulas
  dashboardSheet.getDataRange().recalculate();
  
  SpreadsheetApp.getUi().alert('📊 Dashboard refreshed successfully!');
}
