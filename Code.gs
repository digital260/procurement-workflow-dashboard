/**
 * PROCUREMENT WORKFLOW DASHBOARD - v4.0
 * Features:
 * - Auto-generates PO on manager approval
 * - Auto-generates GRN on quality check PASS
 * - Auto-populates Vendor Master sheet
 * - Complete workflow tracking
 */

// ============================================================================
// INITIALIZATION & MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Procurement Workflow')
    .addItem('📋 Start New Procurement', 'openProcurementForm')
    .addItem('👤 Manager Approval & PO', 'openApprovalForm')
    .addItem('📦 Track Material', 'openTrackingForm')
    .addItem('✅ Quality Check & GRN', 'openQualityForm')
    .addItem('💳 Process Payment', 'openPaymentForm')
    .addItem('🏭 Vendor Master', 'openVendorForm')
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🔧 Setup Sheets', 'initializeSheets')
    .addToUi();
}

// ============================================================================
// MANAGER APPROVAL & PO GENERATION
// ============================================================================

function openApprovalForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data = sheet.getDataRange().getValues();
  
  const pendingRequests = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Pending') {
      pendingRequests.push({
        requestId: data[i][0],
        date: data[i][1],
        requestor: data[i][2],
        amount: data[i][4],
        description: data[i][5],
        email: data[i][1],
        rowIndex: i + 1
      });
    }
  }
  
  let optionsHtml = '<option value="">Select a Request to Approve</option>';
  pendingRequests.forEach(req => {
    optionsHtml += `<option value="${req.rowIndex}|${req.requestId}|${req.requestor}|${req.amount}|${req.email}|${req.description}">
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
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1565c0; margin-bottom: 10px; font-size: 26px; }
        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .tab-button { padding: 10px 20px; background: #f0f0f0; border: none; cursor: pointer; border-radius: 4px; font-weight: 600; }
        .tab-button.active { background: #1565c0; color: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
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
        .po-section { background: #f0f8ff; padding: 15px; border-radius: 4px; margin-top: 15px; border-left: 4px solid #2196F3; }
        .po-section h3 { color: #2196F3; margin-bottom: 10px; }
        .vendor-note { background: #fff8e1; padding: 10px; border-radius: 4px; margin-bottom: 15px; font-size: 13px; color: #ff6f00; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>👤 Manager Approval & PO Portal</h1>
        <p class="subtitle">Review requests, approve/reject, and auto-generate PO</p>
        
        <div class="info">
          <strong>ℹ️ Instructions:</strong> Select a pending request, review details, then either approve (to auto-generate PO) or reject with comments.
        </div>
        
        <div class="tabs">
          <button class="tab-button active" onclick="switchTab('approval')">👤 Approval</button>
          <button class="tab-button" onclick="switchTab('status')">📊 Status</button>
        </div>
        
        <!-- APPROVAL TAB -->
        <div id="approval" class="tab-content active">
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
              
              <!-- PO DETAILS SECTION -->
              <div id="poDetailsSection" style="display: none;" class="po-section">
                <h3>📄 PO Details (Auto-Generated)</h3>
                
                <div class="vendor-note">
                  💡 <strong>Tip:</strong> Enter new vendor details below. If vendor exists, details will auto-populate when you use it in future.
                </div>
                
                <div class="form-group">
                  <label for="vendorName">Vendor Name *</label>
                  <input type="text" id="vendorName" placeholder="Enter vendor name" required>
                </div>
                
                <div class="form-group">
                  <label for="vendorContact">Vendor Contact Email/Phone</label>
                  <input type="text" id="vendorContact" placeholder="contact@vendor.com or 9876543210">
                </div>
                
                <div class="form-group">
                  <label for="items">Items Description *</label>
                  <textarea id="items" placeholder="List items to be purchased" required></textarea>
                </div>
                
                <div class="form-group">
                  <label for="deliveryDate">Delivery Date *</label>
                  <input type="date" id="deliveryDate" required>
                </div>
              </div>
              
              <div class="button-group">
                <button type="button" class="approve" onclick="submitApproval('APPROVED')">✅ Approve & Generate PO</button>
                <button type="button" class="reject" onclick="submitApproval('REJECTED')">❌ Reject Request</button>
              </div>
              
              <div class="success" id="successMsg">✅ Request processed and PO generated successfully!</div>
              <div class="error" id="errorMsg"></div>
            </div>
          </form>
        </div>
        
        <!-- STATUS TAB -->
        <div id="status" class="tab-content">
          <h3>📊 Request Status Summary</h3>
          <div class="details">
            <div class="detail-row">
              <span class="detail-label">Total Pending:</span>
              <span class="detail-value"><strong>${pendingRequests.length}</strong></span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Total Amount Pending:</span>
              <span class="detail-value"><strong>₹${pendingRequests.reduce((sum, req) => sum + parseFloat(req.amount), 0).toLocaleString()}</strong></span>
            </div>
          </div>
        </div>
      </div>
      
      <script>
        function switchTab(tab) {
          document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
          document.querySelectorAll('.tab-button').forEach(el => el.classList.remove('active'));
          document.getElementById(tab).classList.add('active');
          event.target.classList.add('active');
        }
        
        function loadRequestDetails() {
          const selectElement = document.getElementById('request');
          const selectedValue = selectElement.value;
          
          if (!selectedValue) {
            document.getElementById('detailsSection').style.display = 'none';
            return;
          }
          
          const [rowIndex, requestId, requestor, amount, email, description] = selectedValue.split('|');
          
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
            <div class="detail-row">
              <span class="detail-label">Description:</span>
              <span class="detail-value">\${description}</span>
            </div>
          \`;
          
          document.getElementById('requestDetails').innerHTML = detailsHtml;
          document.getElementById('detailsSection').style.display = 'block';
          document.getElementById('poDetailsSection').style.display = 'block';
        }
        
        function submitApproval(status) {
          const selectElement = document.getElementById('request');
          const selectedValue = selectElement.value;
          
          if (!selectedValue) {
            document.getElementById('errorMsg').textContent = 'Please select a request';
            document.getElementById('errorMsg').style.display = 'block';
            return;
          }
          
          const [rowIndex, requestId, requestor, amount, email, description] = selectedValue.split('|');
          
          if (status === 'APPROVED' && (!document.getElementById('vendorName').value || !document.getElementById('items').value)) {
            document.getElementById('errorMsg').textContent = 'Please fill in Vendor Name and Items';
            document.getElementById('errorMsg').style.display = 'block';
            return;
          }
          
          const data = {
            rowIndex: parseInt(rowIndex),
            requestId: requestId,
            status: status,
            approverName: document.getElementById('approverName').value,
            comments: document.getElementById('comments').value,
            approvalDate: new Date().toLocaleString(),
            email: email,
            requestor: requestor,
            amount: amount,
            vendorName: document.getElementById('vendorName').value,
            vendorContact: document.getElementById('vendorContact').value,
            items: document.getElementById('items').value,
            deliveryDate: document.getElementById('deliveryDate').value
          };
          
          if (!data.approverName) {
            document.getElementById('errorMsg').textContent = 'Please enter approver name';
            document.getElementById('errorMsg').style.display = 'block';
            return;
          }
          
          google.script.run.withSuccessHandler(function(result) {
            document.getElementById('successMsg').style.display = 'block';
            document.getElementById('successMsg').innerHTML = result.message;
            setTimeout(() => google.script.host.close(), 2500);
          }).withFailureHandler(function(error) {
            document.getElementById('errorMsg').textContent = 'Error: ' + error;
            document.getElementById('errorMsg').style.display = 'block';
          }).submitManagerApproval(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '👤 Manager Approval & PO');
}

function submitManagerApproval(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName('Requests');
  
  // Update Requests sheet
  requestsSheet.getRange(data.rowIndex, 4).setValue(data.status);
  requestsSheet.getRange(data.rowIndex, 7).setValue(data.approverName);
  requestsSheet.getRange(data.rowIndex, 8).setValue(new Date());
  requestsSheet.getRange(data.rowIndex, 9).setValue(data.comments);
  
  let resultMessage = '';
  
  if (data.status === 'APPROVED') {
    // Generate PO Number
    const poNumber = generatePONumber();
    
    // Add to PO Master sheet
    const poSheet = ss.getSheetByName('PO Master');
    const poRow = [
      poNumber,
      data.requestId,
      data.vendorName,
      data.vendorContact,
      new Date(),
      data.amount,
      data.items,
      data.deliveryDate,
      'Active',
      data.approverName
    ];
    poSheet.appendRow(poRow);
    
    // Update Requests sheet with PO Number
    requestsSheet.getRange(data.rowIndex, 10).setValue(poNumber);
    
    // AUTO-ADD VENDOR TO VENDOR MASTER
    addVendorToMaster(data.vendorName, data.vendorContact);
    
    resultMessage = `✅ Request APPROVED!\n📄 PO Generated: <strong>${poNumber}</strong>\n��� Vendor added to Master`;
    
    // Send approval email
    const subject = `✅ Your Procurement Approved - PO ${poNumber}`;
    const message = `
Dear ${data.requestor},

Your procurement request has been APPROVED by the manager.

REQUEST DETAILS:
- Request ID: ${data.requestId}
- Amount: ₹${data.amount}
- Status: APPROVED

PURCHASE ORDER GENERATED:
- PO Number: ${poNumber}
- Vendor: ${data.vendorName}
- Delivery Date: ${data.deliveryDate}
- Items: ${data.items}

${data.comments ? `Manager Comments:\n${data.comments}\n` : ''}

You can now track your material using the PO Number: ${poNumber}

Best regards,
Procurement Team
    `;
    
    if (data.email && data.email.includes('@')) {
      GmailApp.sendEmail(data.email, subject, message);
    }
  } else {
    resultMessage = `❌ Request REJECTED by Manager`;
    
    // Send rejection email
    const subject = `❌ Your Procurement Request Rejected - ${data.requestId}`;
    const message = `
Dear ${data.requestor},

Your procurement request has been REJECTED by the manager.

REQUEST DETAILS:
- Request ID: ${data.requestId}
- Amount: ₹${data.amount}
- Status: REJECTED
- Rejected By: ${data.approverName}

${data.comments ? `Manager Comments:\n${data.comments}` : 'No comments provided'}

Please contact your manager for more information.

Best regards,
Procurement Team
    `;
    
    if (data.email && data.email.includes('@')) {
      GmailApp.sendEmail(data.email, subject, message);
    }
  }
  
  Logger.log(`Request ${data.requestId} - ${data.status} by ${data.approverName}`);
  
  return { message: resultMessage };
}

// Add vendor to Vendor Master sheet
function addVendorToMaster(vendorName, vendorContact) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor Master');
  const vendorData = vendorSheet.getDataRange().getValues();
  
  // Check if vendor already exists
  for (let i = 1; i < vendorData.length; i++) {
    if (vendorData[i][1] && vendorData[i][1].toLowerCase() === vendorName.toLowerCase()) {
      Logger.log('Vendor already exists: ' + vendorName);
      return; // Vendor already exists
    }
  }
  
  // Add new vendor
  const vendorId = 'VEN-' + String(vendorData.length).padStart(5, '0');
  const row = [
    vendorId,
    vendorName,
    '',
    vendorContact,
    '',
    '',
    '',
    '',
    '',
    'Yes'
  ];
  
  vendorSheet.appendRow(row);
  Logger.log('Vendor added: ' + vendorName);
}

// Generate unique PO Number
function generatePONumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName('PO Master');
  const allData = poSheet.getDataRange().getValues();
  
  let maxNumber = 0;
  for (let i = 1; i < allData.length; i++) {
    const poNum = allData[i][0];
    if (poNum && poNum.startsWith('PO-')) {
      const num = parseInt(poNum.replace('PO-', ''));
      if (num > maxNumber) maxNumber = num;
    }
  }
  
  return 'PO-' + String(maxNumber + 1).padStart(5, '0');
}

// ============================================================================
// MATERIAL TRACKING
// ============================================================================

function openTrackingForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName('PO Master');
  const poData = poSheet.getDataRange().getValues();
  
  let poOptions = '<option value="">Select a PO Number</option>';
  for (let i = 1; i < poData.length; i++) {
    if (poData[i][0]) {
      poOptions += `<option value="${poData[i][0]}|${poData[i][2]}|${poData[i][5]}">
        ${poData[i][0]} - ${poData[i][2]} (₹${poData[i][5]})
      </option>`;
    }
  }
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #FF9800; margin-bottom: 10px; }
        .subtitle { color: #666; margin-bottom: 20px; }
        .info { background: #fff3cd; padding: 12px; border-radius: 4px; margin-bottom: 20px; color: #856404; border: 1px solid #ffc107; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], select { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        button { width: 100%; padding: 12px; background: #FF9800; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 20px; }
        button:hover { background: #e68900; }
        .po-details { background: #f0f8ff; padding: 12px; border-radius: 4px; margin-bottom: 15px; border-left: 4px solid #FF9800; display: none; }
        .detail-row { display: flex; justify-content: space-between; margin-bottom: 8px; }
        .detail-label { font-weight: 600; color: #333; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>📦 Track Material Transit</h1>
        <p class="subtitle">Update shipment status and tracking</p>
        
        <div class="info">
          <strong>ℹ️ Note:</strong> Select an approved PO to track material. PO number will be auto-filled.
        </div>
        
        <form>
          <div class="form-group">
            <label for="poNumber">Select PO Number *</label>
            <select id="poNumber" required onchange="loadPODetails()">
              ${poOptions}
            </select>
          </div>
          
          <div id="poDetailsDiv" class="po-details">
            <div class="detail-row">
              <span class="detail-label">Vendor:</span>
              <span id="vendorName"></span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Amount:</span>
              <span id="poAmount"></span>
            </div>
          </div>
          
          <div class="form-group">
            <label for="trackingNumber">Tracking Number (from Courier) *</label>
            <input type="text" id="trackingNumber" placeholder="e.g., TRK123456789" required>
          </div>
          
          <div class="form-group">
            <label for="courier">Courier Name</label>
            <input type="text" id="courier" placeholder="e.g., DHL, FedEx, Local Courier">
          </div>
          
          <div class="form-group">
            <label for="status">Status *</label>
            <select id="status" required>
              <option>Select Status</option>
              <option>Dispatched</option>
              <option>In Transit</option>
              <option>Out for Delivery</option>
              <option>Delivered</option>
              <option>Delayed</option>
            </select>
          </div>
          
          <div class="form-group">
            <label for="location">Current Location</label>
            <input type="text" id="location" placeholder="e.g., Delhi Warehouse, Mumbai">
          </div>
          
          <div class="form-group">
            <label for="expectedDelivery">Expected Delivery Date</label>
            <input type="date" id="expectedDelivery">
          </div>
          
          <button type="button" onclick="submitTracking()">📤 Update Transit Status</button>
        </form>
      </div>
      
      <script>
        function loadPODetails() {
          const poSelect = document.getElementById('poNumber');
          const selectedValue = poSelect.value;
          
          if (!selectedValue) {
            document.getElementById('poDetailsDiv').style.display = 'none';
            return;
          }
          
          const [poNum, vendor, amount] = selectedValue.split('|');
          document.getElementById('vendorName').textContent = vendor;
          document.getElementById('poAmount').textContent = '₹' + amount;
          document.getElementById('poDetailsDiv').style.display = 'block';
        }
        
        function submitTracking() {
          const poNumber = document.getElementById('poNumber').value;
          
          if (!poNumber) {
            alert('❌ Please select a PO number');
            return;
          }
          
          const data = {
            poNumber: poNumber,
            trackingNumber: document.getElementById('trackingNumber').value,
            courier: document.getElementById('courier').value,
            status: document.getElementById('status').value,
            location: document.getElementById('location').value,
            expectedDelivery: document.getElementById('expectedDelivery').value
          };
          
          if (!data.trackingNumber || !data.status) {
            alert('❌ Please fill in Tracking Number and Status');
            return;
          }
          
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
    data.courier,
    data.location,
    ''
  ];
  
  sheet.appendRow(row);
  Logger.log('Transit tracking updated: ' + transitId + ' for PO: ' + data.poNumber);
}

// ============================================================================
// QUALITY CHECK & GRN GENERATION
// ============================================================================

function openQualityForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transitSheet = ss.getSheetByName('Material Transit');
  const transitData = transitSheet.getDataRange().getValues();
  
  let poOptions = '<option value="">Select a PO Number</option>';
  for (let i = 1; i < transitData.length; i++) {
    if (transitData[i][1]) {
      poOptions += `<option value="${transitData[i][1]}">
        ${transitData[i][1]} - Tracking: ${transitData[i][2]}
      </option>`;
    }
  }
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #9C27B0; margin-bottom: 10px; font-size: 24px; }
        .subtitle { color: #666; margin-bottom: 20px; }
        .info { background: #f3e5f5; padding: 12px; border-radius: 4px; margin-bottom: 20px; color: #7b1fa2; border-left: 4px solid #9C27B0; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], select, textarea { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        textarea { min-height: 80px; }
        .button-group { display: flex; gap: 10px; margin-top: 20px; }
        button { 
          flex: 1;
          padding: 12px; 
          border: none; 
          border-radius: 4px; 
          cursor: pointer; 
          font-weight: bold;
          color: white;
        }
        .pass { background: #4CAF50; }
        .pass:hover { background: #45a049; }
        .fail { background: #f44336; }
        .fail:hover { background: #da190b; }
        .grn-section { background: #e8f5e9; padding: 15px; border-radius: 4px; margin-top: 15px; border-left: 4px solid #4CAF50; display: none; }
        .grn-section h3 { color: #2e7d32; margin-bottom: 10px; }
        .detail-row { display: flex; justify-content: space-between; margin-bottom: 8px; }
        .detail-label { font-weight: 600; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>✅ Quality Check & GRN</h1>
        <p class="subtitle">Inspect material and auto-generate GRN if quality passes</p>
        
        <div class="info">
          <strong>ℹ️ Note:</strong> Select PASS to auto-generate GRN (Goods Receipt Note). Select FAIL to raise debit note.
        </div>
        
        <form id="qualityForm">
          <div class="form-group">
            <label for="poNumber">PO Number *</label>
            <select id="poNumber" required onchange="loadPODetails()">
              ${poOptions}
            </select>
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
          
          <!-- GRN SECTION (shown when PASS selected) -->
          <div id="grnSection" class="grn-section">
            <h3>📄 GRN Auto-Generation</h3>
            <div class="detail-row">
              <span class="detail-label">Status:</span>
              <span><strong>GRN will be auto-generated on PASS</strong></span>
            </div>
            <p style="font-size: 13px; color: #666; margin-top: 8px;">GRN (Goods Receipt Note) number will be created automatically and linked to this PO.</p>
          </div>
          
          <div class="button-group">
            <button type="button" class="pass" onclick="submitQuality('PASS')">✅ PASS & Generate GRN</button>
            <button type="button" class="fail" onclick="submitQuality('FAIL')">❌ FAIL & Raise Debit Note</button>
          </div>
        </form>
      </div>
      
      <script>
        document.getElementById('qualityStatus').addEventListener('change', function() {
          if (this.value === 'PASS') {
            document.getElementById('grnSection').style.display = 'block';
          } else {
            document.getElementById('grnSection').style.display = 'none';
          }
        });
        
        function loadPODetails() {
          // Can add additional details loading here
        }
        
        function submitQuality(decision) {
          const data = {
            poNumber: document.getElementById('poNumber').value,
            qualityStatus: document.getElementById('qualityStatus').value,
            defects: document.getElementById('defects').value,
            checkedBy: document.getElementById('checkedBy').value,
            notes: document.getElementById('notes').value,
            decision: decision
          };
          
          if (!data.poNumber || !data.qualityStatus || !data.checkedBy) {
            alert('❌ Please fill in all required fields');
            return;
          }
          
          google.script.run.withSuccessHandler(function(result) {
            alert(result.message);
            google.script.host.close();
          }).submitQualityCheck(data);
        }
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '✅ Quality Check & GRN');
}

function submitQualityCheck(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qualitySheet = ss.getSheetByName('Quality Check');
  
  const checkId = 'QC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  let grnNumber = '';
  
  if (data.decision === 'PASS') {
    // Generate GRN Number
    grnNumber = generateGRNNumber();
    
    // Add GRN to Quality Check sheet
    const row = [
      checkId,
      data.poNumber,
      new Date(),
      data.qualityStatus,
      data.defects,
      data.checkedBy,
      new Date(),
      'Proceed to Payment',
      grnNumber,
      data.notes
    ];
    qualitySheet.appendRow(row);
    
    // Create entry in a separate GRN tracking (using Quality Check column)
    Logger.log('Quality check passed and GRN generated: ' + grnNumber + ' for PO: ' + data.poNumber);
    
    return { message: `✅ Quality PASS!\n📄 GRN Generated: ${grnNumber}\n✓ Proceed to Payment` };
  } else {
    // FAIL - Raise Debit Note
    const debitNoteNumber = 'DBIT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    
    const row = [
      checkId,
      data.poNumber,
      new Date(),
      data.qualityStatus,
      data.defects,
      data.checkedBy,
      new Date(),
      'Raise Debit Note: ' + debitNoteNumber,
      'PENDING',
      data.notes
    ];
    qualitySheet.appendRow(row);
    
    Logger.log('Quality check failed. Debit note: ' + debitNoteNumber + ' for PO: ' + data.poNumber);
    
    return { message: `❌ Quality FAIL!\n📝 Debit Note: ${debitNoteNumber}\n⚠️ Contact Vendor` };
  }
}

// Generate unique GRN Number
function generateGRNNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qualitySheet = ss.getSheetByName('Quality Check');
  const allData = qualitySheet.getDataRange().getValues();
  
  let maxNumber = 0;
  for (let i = 1; i < allData.length; i++) {
    const grnNum = allData[i][8]; // GRN column
    if (grnNum && grnNum.startsWith('GRN-')) {
      const num = parseInt(grnNum.replace('GRN-', ''));
      if (num > maxNumber) maxNumber = num;
    }
  }
  
  return 'GRN-' + String(maxNumber + 1).padStart(5, '0');
}

// ============================================================================
// PAYMENT PROCESSING
// ============================================================================

function openPaymentForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName('PO Master');
  const poData = poSheet.getDataRange().getValues();
  
  let poOptions = '<option value="">Select a PO Number</option>';
  for (let i = 1; i < poData.length; i++) {
    if (poData[i][0]) {
      poOptions += `<option value="${poData[i][0]}|${poData[i][2]}">
        ${poData[i][0]} - ${poData[i][2]}
      </option>`;
    }
  }
  
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
          <select id="poNumber" required onchange="updateVendor()">
            ${poOptions}
          </select>
        </div>
        
        <div class="form-group">
          <label for="vendorName">Vendor Name</label>
          <input type="text" id="vendorName" placeholder="Auto-filled" readonly>
        </div>
        
        <div class="form-group">
          <label for="invoiceNumber">Invoice Number *</label>
          <input type="text" id="invoiceNumber" placeholder="Enter invoice number" required>
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
        function updateVendor() {
          const poSelect = document.getElementById('poNumber');
          const selectedValue = poSelect.value;
          
          if (selectedValue) {
            const [poNum, vendor] = selectedValue.split('|');
            document.getElementById('vendorName').value = vendor;
          }
        }
        
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
          
          if (!data.poNumber || !data.invoiceNumber || !data.amount) {
            alert('❌ Please fill in all required fields');
            return;
          }
          
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

// ============================================================================
// VENDOR MASTER MANAGEMENT
// ============================================================================

function openVendorForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor Master');
  const vendorData = vendorSheet.getDataRange().getValues();
  
  let vendorOptions = '<option value="">Select a Vendor</option>';
  for (let i = 1; i < vendorData.length; i++) {
    if (vendorData[i][1]) {
      vendorOptions += `<option value="${vendorData[i][0]}|${vendorData[i][1]}|${vendorData[i][2]}|${vendorData[i][3]}|${vendorData[i][4]}|${vendorData[i][5]}|${vendorData[i][6]}|${vendorData[i][7]}|${vendorData[i][8]}">
        ${vendorData[i][1]} - ${vendorData[i][3]}
      </option>`;
    }
  }
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #00BCD4; margin-bottom: 10px; font-size: 24px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .tab-button { padding: 10px 20px; background: #f0f0f0; border: none; cursor: pointer; border-radius: 4px; font-weight: 600; }
        .tab-button.active { background: #00BCD4; color: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        input[type="text"], input[type="email"], select, textarea { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px; }
        button { width: 100%; padding: 12px; background: #00BCD4; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 20px; }
        button:hover { background: #0097a7; }
        .info { background: #e0f2f1; padding: 12px; border-radius: 4px; margin-bottom: 20px; color: #00695c; border-left: 4px solid #00BCD4; }
        .vendor-list { max-height: 400px; overflow-y: auto; }
        .vendor-item { background: #f9f9f9; padding: 12px; margin-bottom: 10px; border-radius: 4px; border-left: 4px solid #00BCD4; }
        .vendor-item strong { color: #00695c; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🏭 Vendor Master Management</h1>
        
        <div class="tabs">
          <button class="tab-button active" onclick="switchTab('view')">📋 View Vendors</button>
          <button class="tab-button" onclick="switchTab('add')">➕ Add New Vendor</button>
        </div>
        
        <!-- VIEW TAB -->
        <div id="view" class="tab-content active">
          <div class="info">
            <strong>ℹ️ Note:</strong> All vendors added during PO creation appear automatically here.
          </div>
          
          <div class="vendor-list" id="vendorList">
            <div class="vendor-item">Loading vendors...</div>
          </div>
        </div>
        
        <!-- ADD TAB -->
        <div id="add" class="tab-content">
          <form id="vendorForm">
            <div class="form-group">
              <label for="vendorName">Vendor Name *</label>
              <input type="text" id="vendorName" placeholder="Company name" required>
            </div>
            
            <div class="form-group">
              <label for="contactPerson">Contact Person</label>
              <input type="text" id="contactPerson" placeholder="Name of contact person">
            </div>
            
            <div class="form-group">
              <label for="email">Email *</label>
              <input type="email" id="email" placeholder="vendor@company.com" required>
            </div>
            
            <div class="form-group">
              <label for="phone">Phone</label>
              <input type="text" id="phone" placeholder="9876543210">
            </div>
            
            <div class="form-group">
              <label for="address">Address</label>
              <textarea id="address" placeholder="Full address" style="min-height: 60px;"></textarea>
            </div>
            
            <div class="form-group">
              <label for="city">City</label>
              <input type="text" id="city" placeholder="City">
            </div>
            
            <div class="form-group">
              <label for="bankDetails">Bank Details</label>
              <textarea id="bankDetails" placeholder="Account number, IFSC code, etc." style="min-height: 60px;"></textarea>
            </div>
            
            <div class="form-group">
              <label for="paymentTerms">Payment Terms</label>
              <input type="text" id="paymentTerms" placeholder="e.g., Net 30">
            </div>
            
            <button type="button" onclick="addNewVendor()">➕ Add Vendor</button>
          </form>
        </div>
      </div>
      
      <script>
        function switchTab(tab) {
          document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
          document.querySelectorAll('.tab-button').forEach(el => el.classList.remove('active'));
          document.getElementById(tab).classList.add('active');
          event.target.classList.add('active');
          
          if (tab === 'view') {
            loadVendorList();
          }
        }
        
        function loadVendorList() {
          google.script.run.getVendorList(function(vendors) {
            let html = '';
            if (vendors.length === 0) {
              html = '<div class="vendor-item">No vendors found</div>';
            } else {
              vendors.forEach(v => {
                html += \`
                  <div class="vendor-item">
                    <strong>\${v.name}</strong><br>
                    📧 \${v.email}<br>
                    📱 \${v.phone || 'N/A'}<br>
                    📍 \${v.city || 'N/A'}
                  </div>
                \`;
              });
            }
            document.getElementById('vendorList').innerHTML = html;
          });
        }
        
        function addNewVendor() {
          const data = {
            vendorName: document.getElementById('vendorName').value,
            contactPerson: document.getElementById('contactPerson').value,
            email: document.getElementById('email').value,
            phone: document.getElementById('phone').value,
            address: document.getElementById('address').value,
            city: document.getElementById('city').value,
            bankDetails: document.getElementById('bankDetails').value,
            paymentTerms: document.getElementById('paymentTerms').value
          };
          
          if (!data.vendorName || !data.email) {
            alert('❌ Please fill in Vendor Name and Email');
            return;
          }
          
          google.script.run.withSuccessHandler(function() {
            alert('✅ Vendor added successfully!');
            document.getElementById('vendorForm').reset();
            google.script.host.close();
          }).addVendorFromForm(data);
        }
        
        // Load vendors on page open
        window.addEventListener('load', function() {
          loadVendorList();
        });
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '🏭 Vendor Master');
}

function getVendorList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor Master');
  const vendorData = vendorSheet.getDataRange().getValues();
  
  let vendors = [];
  for (let i = 1; i < vendorData.length; i++) {
    if (vendorData[i][1]) {
      vendors.push({
        id: vendorData[i][0],
        name: vendorData[i][1],
        contact: vendorData[i][2],
        email: vendorData[i][3],
        phone: vendorData[i][4],
        address: vendorData[i][5],
        city: vendorData[i][6],
        bank: vendorData[i][7],
        terms: vendorData[i][8]
      });
    }
  }
  
  return vendors;
}

function addVendorFromForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName('Vendor Master');
  const vendorData = vendorSheet.getDataRange().getValues();
  
  // Check if vendor already exists
  for (let i = 1; i < vendorData.length; i++) {
    if (vendorData[i][1] && vendorData[i][1].toLowerCase() === data.vendorName.toLowerCase()) {
      throw new Error('Vendor already exists!');
    }
  }
  
  // Add new vendor
  const vendorId = 'VEN-' + String(vendorData.length).padStart(5, '0');
  const row = [
    vendorId,
    data.vendorName,
    data.contactPerson,
    data.email,
    data.phone,
    data.address,
    data.city,
    data.bankDetails,
    data.paymentTerms,
    'Yes'
  ];
  
  vendorSheet.appendRow(row);
  Logger.log('Vendor added: ' + data.vendorName);
}

// ============================================================================
// INITIALIZATION & DASHBOARD
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
  
  sheet.appendRow(['PROCUREMENT WORKFLOW DASHBOARD', '', new Date()]);
  sheet.appendRow(['']);
  sheet.appendRow(['Metric', 'Count', 'Status']);
  sheet.appendRow(['Total Requests', '=COUNTA(Requests!A2:A)', '']);
  sheet.appendRow(['Pending Approval', '=COUNTIF(Requests!D2:D,"Pending")', '']);
  sheet.appendRow(['Approved Requests', '=COUNTIF(Requests!D2:D,"APPROVED")', '']);
  sheet.appendRow(['Total POs Generated', '=COUNTA(\'PO Master\'!A2:A)', '']);
  sheet.appendRow(['In Transit', '=COUNTIF(\'Material Transit\'!E2:E,"In Transit")', '']);
  sheet.appendRow(['GRN Generated', '=COUNTIF(\'Quality Check\'!I2:I,"GRN-*")', '']);
  sheet.appendRow(['Quality Passed', '=COUNTIF(\'Quality Check\'!D2:D,"PASS")', '']);
  sheet.appendRow(['Payments Processed', '=COUNTA(Payment!A2:A)', '']);
  sheet.appendRow(['Total Vendors', '=COUNTA(\'Vendor Master\'!A2:A)', '']);
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
    'Comments',
    'PO Number'
  ];
  
  sheet.appendRow(headers);
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  
  for (let i = 0; i < headers.length; i++) {
    sheet.setColumnWidth(i + 1, 120);
  }
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
  
  for (let i = 0; i < headers.length; i++) {
    sheet.setColumnWidth(i + 1, 120);
  }
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
    'GRN Number',
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

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  
  dashboardSheet.getDataRange().recalculate();
  
  SpreadsheetApp.getUi().alert('📊 Dashboard refreshed successfully!');
}

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
    data.notes,
    ''
  ];
  
  sheet.appendRow(row);
  
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

Next Steps:
1. Your request will be reviewed by the manager
2. If approved, a PO will be auto-generated with a unique PO Number
3. You will receive the PO Number via email
4. Use that PO Number to track your material

You will receive updates as the request progresses through the workflow.

Best regards,
Procurement Team
  `;
  
  if (mailTo && mailTo.includes('@')) {
    GmailApp.sendEmail(mailTo, subject, message);
  }
  
  Logger.log('Request submitted: ' + requestId);
}

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
            <label for="deliveryDate">Required Delivery Date *</label>
            <input type="date" id="deliveryDate" required>
          </div>
          
          <div class="form-group">
            <label for="notes">Additional Notes</label>
            <textarea id="notes" placeholder="Any special requirements"></textarea>
          </div>
          
          <button type="submit">📤 Submit Request</button>
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
            deliveryDate: document.getElementById('deliveryDate').value,
            notes: document.getElementById('notes').value
          };
          
          google.script.run.withSuccessHandler(function() {
            alert('✅ Request submitted successfully!');
            google.script.host.close();
          }).submitProcurementRequest(data);
        });
      </script>
    </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '🛒 New Procurement Request');
}
