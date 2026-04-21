/**
 * PROCUREMENT WORKFLOW DASHBOARD - WITH MANAGER APPROVAL
 * Complete implementation for Google Sheets
 * Author: Procurement Team
 * Version: 2.0
 */

// ============================================================================
// INITIALIZATION & MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Procurement Workflow')
    .addItem('📋 Start New Procurement', 'openProcurementForm')
    .addItem('👤 Manager Approval', 'openApprovalForm')
    .addItem('📦 Track Material', 'openTrackingForm')
    .addItem('✅ Quality Check', 'openQualityForm')
    .addItem('💳 Process Payment', 'openPaymentForm')
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🔧 Setup Sheets', 'initializeSheets')
    .addToUi();
}

// ============================================================================
// MANAGER APPROVAL SECTION
// ============================================================================

function openApprovalForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data = sheet.getDataRange().getValues();
  
  // Get pending requests
  const pendingRequests = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Pending') {
      pendingRequests.push({
        requestId: data[i][0],
        date: data[i][1],
        requestor: data[i][2],
        amount: data[i][4],
        description: data[i][5],
        rowIndex: i + 1
      });
    }
  }
  
  let optionsHtml = '<option value="">Select a Request to Approve</option>';
  pendingRequests.forEach(req => {
    optionsHtml += `<option value="${req.rowIndex}|${req.requestId}|${req.requestor}|${req.amount}">
      ${req.requestId} - ${req.requestor} - ₹${req.amount}
    </option>`;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1565c0; margin-bottom: 10px; font-size: 26px; }
        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; color: #333; }
        input[type="text"], select, textarea {
          width: 100%;
          padding: 12px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 14px;
        }
        textarea { resize: vertical; min-height: 80px; }
        .button-group { display: flex; gap: 10px; margin-top: 20px; }
        button { 
          flex: 1;
          padding: 12px; 
          color: white; 
          border: none; 
          border-radius: 4px; 
          font-size: 16px; 
          font-weight: bold;
          cursor: pointer;
        }
        .approve { background: #4CAF50; }
        .approve:hover { background: #45a049; }
        .reject { background: #f44336; }
        .reject:hover { background: #da190b; }
        .info { background: #e3f2fd; padding: 15px; border-radius: 4px; margin-bottom: 20px; color: #1565c0; font-size: 13px; border-left: 4px solid #1565c0; }
        .details { background: #f9f9f9; padding: 15px; border-radius: 4px; margin-bottom: 20px; border: 1px solid #ddd; }
        .detail-row { display: flex; justify-content: space-between; margin-bottom: 10px; }
        .detail-label { font-weight: 600; color: #333; }
        .detail-value { color: #666; }
        .success { color: #4CAF50; display: none; margin-top: 10px; text-align: center; }
        .error { color: #f44336; display: none; margin-top: 10px; }
        .warning { background: #fff3cd; padding: 12px; border-radius: 4px; color: #856404; margin-bottom: 20px; border: 1px solid #ffc107; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>👤 Manager Approval Portal</h1>
        <p class="subtitle">Review and approve/reject procurement requests</p>
        
        <div class="info">
          <strong>ℹ️ Instructions:</strong> Select a pending request below and either approve it to proceed to PO creation or reject it with comments.
        </div>
        
        ${pendingRequests.length === 0 ? 
          `<div class="warning">✓ No pending requests to approve!</div>` 
          : ''}
        
        <form id="approvalForm">
          <div class="form-group">
            <label for="request">Pending Requests *</label>
            <select id="request" required onchange="loadRequestDetails()">
              ${optionsHtml}
            </select>
          </div>
          
          <div id="detailsSection" style="display: none;">
            <div class="details" id="requestDetails"></div>
            
            <div class="form-group">
              <label for="approverName">Approver Name *</label>
              <input type="text" id="approverName" placeholder="Your name" required>
            </div>
            
            <div class="form-group">
              <label for="comments">Comments/Notes</label>
              <textarea id="comments" placeholder="Add any comments or feedback..."></textarea>
            </div>
            
            <div class="button-group">
              <button type="button" class="approve" onclick="submitApproval('APPROVED')">✅ Approve Request</button>
              <button type="button" class="reject" onclick="submitApproval('REJECTED')">❌ Reject Request</button>
            </div>
            
            <div class="success" id="successMsg">✅ Request processed successfully!</div>
            <div class="error" id="errorMsg"></div>
          </div>
        </form>
      </div>
      
      <script>
        function loadRequestDetails() {
          const selectElement = document.getElementById('request');
          const selectedValue = selectElement.value;
          
          if (!selectedValue) {
            document.getElementById('detailsSection').style.display = 'none';
            return;
          }
          
          const [rowIndex, requestId, requestor, amount] = selectedValue.split('|');
          
          const detailsHtml = \`
            <div class="detail-row">
              <span class="detail-label">Request ID:</span>
              <span class="detail-value"><strong>\${requestId}</strong></span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Requestor:</span>
              <span class="detail-value">\${requestor}</span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Amount:</span>
              <span class="detail-value"><strong>₹\${amount}</strong></span>
            </div>
          \`;
          
          document.getElementById('requestDetails').innerHTML = detailsHtml;
          document.getElementById('detailsSection').style.display = 'block';
        }
        
        function submitApproval(status) {
          const selectElement = document.getElementById('request');
          const selectedValue = selectElement.value;
          
          if (!selectedValue) {
            document.getElementById('errorMsg').textContent = 'Please select a request';
            document.getElementById('errorMsg').style.display = 'block';
            return;
          }
          
          const [rowIndex, requestId, requestor, amount] = selectedValue.split('|');
          
          const data = {
            rowIndex: parseInt(rowIndex),
            requestId: requestId,
            status: status,
            approverName: document.getElementById('approverName').value,
            comments: document.getElementById('comments').value,
            approvalDate: new Date().toLocaleString()
          };
          
          if (!data.approverName) {
            document.getElementById('errorMsg').textContent = 'Please enter approver name';
            document.getElementById('errorMsg').style.display = 'block';
            return;
          }
          
          google.script.run.withSuccessHandler(function() {
            document.getElementById('successMsg').style.display = 'block';
            setTimeout(() => google.script.host.close(), 2000);
          }).withFailureHandler(function(error) {
            document.getElementById('errorMsg').textContent = 'Error: ' + error;
            document.getElementById('errorMsg').style.display = 'block';
          }).submitManagerApproval(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '👤 Manager Approval');
}

function submitManagerApproval(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Requests');
  const range = sheet.getRange(data.rowIndex, 1, 1, 9);
  const rowData = range.getValues()[0];
  
  // Update status
  sheet.getRange(data.rowIndex, 4).setValue(data.status);
  
  // Update manager approval
  sheet.getRange(data.rowIndex, 7).setValue(data.approverName);
  
  // Update approval date
  sheet.getRange(data.rowIndex, 8).setValue(new Date());
  
  // Update notes with comments
  sheet.getRange(data.rowIndex, 9).setValue(data.comments);
  
  // Send approval notification email
  const requestorEmail = rowData[1]; // Email column (index 1)
  const requestor = rowData[2];
  const amount = rowData[4];
  
  const subject = data.status === 'APPROVED' 
    ? `✅ Your Procurement Request Approved - ${data.requestId}`
    : `❌ Your Procurement Request Rejected - ${data.requestId}`;
  
  const message = data.status === 'APPROVED'
    ? `
Dear ${requestor},

Your procurement request has been APPROVED by the manager.

Request Details:
- Request ID: ${data.requestId}
- Amount: ₹${amount}
- Status: APPROVED
- Approved By: ${data.approverName}
- Approval Date: ${data.approvalDate}

${data.comments ? `\nManager Comments:\n${data.comments}` : ''}

Your request will now proceed to the next stage of the procurement process.

Best regards,
Procurement Team
    `
    : `
Dear ${requestor},

Your procurement request has been REJECTED by the manager.

Request Details:
- Request ID: ${data.requestId}
- Amount: ₹${amount}
- Status: REJECTED
- Rejected By: ${data.approverName}
- Rejection Date: ${data.approvalDate}

${data.comments ? `\nManager Comments:\n${data.comments}` : ''}

Please contact your manager for more information.

Best regards,
Procurement Team
    `;
  
  if (requestorEmail && requestorEmail.includes('@')) {
    GmailApp.sendEmail(requestorEmail, subject, message);
  }
  
  Logger.log(`Request ${data.requestId} - ${data.status} by ${data.approverName}`);
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
  
  setupDashboard();
  setupRequestsSheet();
  setupPOMasterSheet();
  setupMaterialTransitSheet();
  setupQualityCheckSheet();
  setupPaymentSheet();
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
  sheet.appendRow(['Approved Requests', '=COUNTIF(Requests!D2:D,"APPROVED")', '']);
  sheet.appendRow(['Rejected Requests', '=COUNTIF(Requests!D2:D,"REJECTED")', '']);
  sheet.appendRow(['In Transit', '=COUNTIF(\'Material Transit\'!E2:E,"In Transit")', '']);
  sheet.appendRow(['Quality Passed', '=COUNTIF(\'Quality Check\'!D2:D,"PASS")', '']);
  sheet.appendRow(['Payments Pending', '=COUNTIF(Payment!F2:F,"Pending")', '']);
  sheet.appendRow(['Total Amount', '=SUM(\'PO Master\'!F2:F)', '']);
  
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
    'Email',
    'Requestor',
    'Status',
    'Amount',
    'Description',
    'Manager Approval',
    'Approval Date',
    'Comments'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
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
    data.email,
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
  
  dashboardSheet.getDataRange().recalculate();
  
  SpreadsheetApp.getUi().alert('📊 Dashboard refreshed successfully!');
}
