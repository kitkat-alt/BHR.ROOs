const SCRIPT_PROP = PropertiesService.getScriptProperties();
const FOLDER_NAME = "JuristicPropertyImages"; 
const SHEET_NAME = "Properties"; 
const DEFAULT_PASSWORD = "7014"; // รหัสผ่านเริ่มต้น

function setup() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = doc.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = doc.insertSheet(SHEET_NAME);
    const headers = [
      "id", "status", "roomNumber", "type", "floor", "size", 
      "rentPrice", "sellPrice", "details", "images", 
      "ownerContract", "tenantContract", 
      "contractStart", "contractEnd", "updatedAt", "adminNote"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  let folder;
  if (folders.hasNext()) { folder = folders.next(); } 
  else {
    folder = DriveApp.createFolder(FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  SCRIPT_PROP.setProperty("FOLDER_ID", folder.getId());

  if (!SCRIPT_PROP.getProperty("ADMIN_PASSWORD")) {
    SCRIPT_PROP.setProperty("ADMIN_PASSWORD", DEFAULT_PASSWORD);
  }
  
  return ContentService.createTextOutput("Setup Complete.");
}

// Function to Force Reset Password
function resetAdminPassword() {
  SCRIPT_PROP.setProperty("ADMIN_PASSWORD", DEFAULT_PASSWORD);
  return "Password has been reset to: " + DEFAULT_PASSWORD;
}

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = doc.getSheetByName(SHEET_NAME);
    
    let action = e.parameter.action;
    let data = {};
    
    if (e.postData && e.postData.contents) {
      const postBody = JSON.parse(e.postData.contents);
      action = postBody.action;
      data = postBody.data;
    }

    // --- GET SETTINGS ---
    if (action === "get_settings") {
        return responseJSON({
            status: "success",
            data: {
                lineUrl: SCRIPT_PROP.getProperty("LINE_URL") || "#",
                mapUrl: SCRIPT_PROP.getProperty("MAP_URL") || "#"
            }
        });
    }

    // --- SAVE SETTINGS ---
    if (action === "save_settings") {
        if (data.newPassword) {
            const currentPass = SCRIPT_PROP.getProperty("ADMIN_PASSWORD");
            if (data.oldPassword !== currentPass) return responseJSON({ status: "error", message: "รหัสผ่านเดิมไม่ถูกต้อง" });
            SCRIPT_PROP.setProperty("ADMIN_PASSWORD", data.newPassword);
        }
        if (data.lineUrl !== undefined) SCRIPT_PROP.setProperty("LINE_URL", data.lineUrl);
        if (data.mapUrl !== undefined) SCRIPT_PROP.setProperty("MAP_URL", data.mapUrl);
        
        return responseJSON({ status: "success", message: "บันทึกการตั้งค่าเรียบร้อย" });
    }

    // --- LOGIN CHECK ---
    if (action === "login") {
      if (data.password === SCRIPT_PROP.getProperty("ADMIN_PASSWORD")) return responseJSON({ status: "success", message: "Login successful" });
      else return responseJSON({ status: "error", message: "รหัสผ่านไม่ถูกต้อง" });
    }

    // --- CHANGE PASSWORD (Legacy support) ---
    if (action === "change_password") { 
       const currentPass = SCRIPT_PROP.getProperty("ADMIN_PASSWORD");
       if (data.oldPassword === currentPass) {
         SCRIPT_PROP.setProperty("ADMIN_PASSWORD", data.newPassword);
         return responseJSON({ status: "success", message: "เปลี่ยนรหัสผ่านสำเร็จ" });
       } else return responseJSON({ status: "error", message: "รหัสผ่านเดิมไม่ถูกต้อง" });
    }

    // --- READ ROOMS ---
    if (action === "read") {
      const mode = e.parameter.mode || "public";
      const rows = sheet.getDataRange().getValues();
      const headers = rows.shift(); 
      const result = [];
      const rowsToDelete = [];
      const now = new Date();

      rows.forEach((row, index) => {
        const item = {};
        headers.forEach((header, i) => item[header] = row[i]);
        
        // Auto Delete Trash > 30 Days
        if (item.status === "trash") {
          const updateDate = new Date(item.updatedAt);
          const diffTime = Math.abs(now - updateDate);
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          if (diffDays > 30) {
            rowsToDelete.push(index + 2);
            return;
          }
        }
        
        if (mode === "public") {
          if (item.status === "post") {
            delete item.ownerContract; delete item.tenantContract; delete item.contractStart; delete item.contractEnd; delete item.adminNote;
            result.push(item);
          }
        } else { result.push(item); }
      });

      for (let i = rowsToDelete.length - 1; i >= 0; i--) sheet.deleteRow(rowsToDelete[i]);
      return responseJSON({ status: "success", data: result });
    } 
    
    // --- SAVE ROOM DATA ---
    else if (action === "save") {
      const mainFolder = DriveApp.getFolderById(SCRIPT_PROP.getProperty("FOLDER_ID"));
      const roomNum = data.roomNumber.toString().trim();
      let roomFolder;
      
      // Check/Create Subfolder
      const subFolders = mainFolder.getFoldersByName(roomNum);
      if (subFolders.hasNext()) { roomFolder = subFolders.next(); }
      else {
          roomFolder = mainFolder.createFolder(roomNum);
          roomFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
      
      const saveFile = (b64) => {
         if(!b64 || !b64.startsWith('data:')) return "";
         const f = roomFolder.createFile(parseBase64(b64)); // Save to roomFolder
         f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         return f.getId(); 
      };
      
      const saveFileUrl = (b64, customName) => { 
         if(!b64 || !b64.startsWith('data:')) return "";
         const blob = parseBase64(b64);
         if(customName) blob.setName(customName); 
         const f = roomFolder.createFile(blob); // Save to roomFolder
         f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         return f.getUrl();
      };

      let imageUrls = [];
      if (data.newImages) data.newImages.forEach(img => { const id = saveFile(img); if(id) imageUrls.push(id); });
      
      const finalImages = JSON.stringify([...(data.existingImages||[]), ...imageUrls]);
      
      // File Naming Logic
      const nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmm");
      const baseName = `${roomNum}`;
      
      const ownerName = `${baseName}_owner_contract_${nowStr}`;
      const tenantName = `${baseName}_tenant_contract_${nowStr}`;

      const ownerUrl = (data.newOwnerContract && data.newOwnerContract.startsWith('data:')) ? saveFileUrl(data.newOwnerContract, ownerName) : (data.existingOwnerContract || "");
      const tenantUrl = (data.newTenantContract && data.newTenantContract.startsWith('data:')) ? saveFileUrl(data.newTenantContract, tenantName) : (data.existingTenantContract || "");

      const rowData = [
        data.id || Utilities.getUuid(), data.status, data.roomNumber, data.type, data.floor, data.size, 
        data.rentPrice, data.sellPrice, data.details, finalImages, ownerUrl, tenantUrl, 
        data.contractStart || "", data.contractEnd || "", new Date().toISOString(), data.adminNote || ""
      ];

      const allIds = sheet.getRange(2, 1, sheet.getLastRow() - 1 || 1, 1).getValues().flat();
      const rowIndex = allIds.indexOf(data.id);

      // Handle Headers update if new column is missing
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers.includes("adminNote")) {
         sheet.getRange(1, headers.length + 1).setValue("adminNote");
      }

      if (rowIndex !== -1) {
          // Check if adminNote column index exists in existing row, if not, handled by append or expanded range
          // But safer to just set the row logic. 
          // Ideally, we get header length to know where to put what, but here we assume fixed structure 
          // or we can allow sheet to just accept the array.
          sheet.getRange(rowIndex + 2, 1, 1, rowData.length).setValues([rowData]); 
      }
      else sheet.appendRow(rowData);

      return responseJSON({ status: "success", message: "Saved successfully" });
    }

    // --- CREATE QUOTATION ---
    else if (action === "create_quotation") {
       const mainFolder = DriveApp.getFolderById(SCRIPT_PROP.getProperty("FOLDER_ID"));
       const roomNum = data.roomNumber.toString().trim();
       let roomFolder;
       const subFolders = mainFolder.getFoldersByName(roomNum);
       if (subFolders.hasNext()) { roomFolder = subFolders.next(); }
       else {
           roomFolder = mainFolder.createFolder(roomNum);
           roomFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       }

       // Generate PDF
       const htmlContent = generateQuotationHTML(data);
       const dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd");
       const fileName = `Booking_${roomNum}_${data.customerName}_${dateStr}.pdf`;
       
       const blob = Utilities.newBlob(htmlContent, MimeType.HTML).setName(fileName);
       // Note: Direct HTML to PDF in GAS can use DriveApp or specialized Document approach.
       // The simplest working method for simple HTML is creating a temporary file and converting.
       // However, DriveApp.createFile(blob.getAs(MimeType.PDF)) works if HTML is simple.
       
       const pdfFile = roomFolder.createFile(blob.getAs(MimeType.PDF));
       pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       
       return responseJSON({ status: "success", pdfUrl: pdfFile.getUrl() });
    }

    // --- LIST QUOTATIONS ---
    else if (action === "list_quotations") {
       const mainFolder = DriveApp.getFolderById(SCRIPT_PROP.getProperty("FOLDER_ID"));
       const roomNum = data.roomNumber.toString().trim();
       let roomFolder;
       const subFolders = mainFolder.getFoldersByName(roomNum);
       // If no folder, return empty
       if (subFolders.hasNext()) { roomFolder = subFolders.next(); }
       else { return responseJSON({ status: "success", data: [] }); }

       const files = roomFolder.getFiles();
       const result = [];
       while (files.hasNext()) {
         const file = files.next();
         const name = file.getName();
         // Filter for PDF or Quotation files
         if (name.startsWith("Quotation_") || name.toLowerCase().endsWith(".pdf")) {
           result.push({
             name: name,
             url: file.getUrl(),
             date: file.getDateCreated(),
             size: file.getSize()
           });
         }
       }
       // Sort by date desc
       result.sort((a, b) => new Date(b.date) - new Date(a.date));
       
       return responseJSON({ status: "success", data: result });
    }

  } catch (e) { return responseJSON({ status: "error", message: e.toString() + "\\nStack: " + e.stack }); } 
  finally { lock.releaseLock(); }
}

function responseJSON(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function parseBase64(b64) { 
    const split = b64.split(','); 
    return Utilities.newBlob(Utilities.base64Decode(split[1]), split[0].split(':')[1].split(';')[0], "upload_" + new Date().getTime()); 
}

function generateQuotationHTML(d) {
  const dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "d MMM yyyy");
  const timeStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "HH:mm");
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        @page { margin: 0; }
        body { margin: 0; padding: 0; font-family: 'Sarabun', sans-serif; background: #fff; color: #1e293b; -webkit-print-color-adjust: exact; }
        .page { padding: 50px; position: relative; }
        
        .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 40px; padding-bottom: 20px; border-bottom: 2px solid #e2e8f0; }
        .brand { font-size: 28px; font-weight: 800; color: #1e293b; letter-spacing: -0.5px; } /* Increased Font */
        .brand span { color: #4338ca; }
        .doc-title { text-align: right; }
        .doc-name { font-size: 28px; font-weight: 800; color: #4338ca; margin: 0; text-transform: uppercase; } /* Increased Font */
        .doc-meta { font-size: 16px; color: #64748b; margin-top: 5px; } /* Increased Font */

        .grid { display: flex; gap: 40px; margin-bottom: 30px; }
        .col { flex: 1; }
        
        .box-title { font-size: 14px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; } /* Increased Font */
        .box { background: #f8fafc; padding: 25px; border-radius: 12px; border: 1px solid #f1f5f9; } /* Increased Padding */
        
        .info-row { display: flex; justify-content: space-between; margin-bottom: 10px; font-size: 16px; } /* Increased Font & Margin */
        .info-row:last-child { margin-bottom: 0; }
        .info-label { color: #64748b; font-weight: 500; }
        .info-val { font-weight: 700; color: #334155; }

        .highlight-box { background: #eff6ff; border-color: #bfdbfe; color: #1e40af; }
        .highlight-box .info-label { color: #60a5fa; }
        .highlight-box .info-val { color: #1e3a8a; }

        .total-section { margin-top: 30px; text-align: right; padding: 20px; }
        .total-label { font-size: 16px; color: #64748b; font-weight: 600; text-transform: uppercase; } /* Increased Font */
        .total-amount { font-size: 42px; font-weight: 900; color: #4338ca; margin-top: 5px; letter-spacing: -1px; } /* Increased Font */

        .terms { margin-top: 40px; border-top: 1px dashed #e2e8f0; padding-top: 20px; font-size: 14px; color: #94a3b8; line-height: 1.6; } /* Increased Font */
        
        .signature-area { display: flex; justify-content: space-between; margin-top: 80px; }
        .sig-block { width: 45%; text-align: center; }
        .sig-line { border-bottom: 1px solid #cbd5e1; height: 40px; margin-bottom: 10px; }
        .sig-name { font-weight: 700; font-size: 16px; color: #334155; } /* Increased Font */
        .sig-role { font-size: 14px; color: #94a3b8; font-weight: 500; } /* Increased Font */
      </style>
    </head>
    <body>
      <div class="page">
        <!-- HEADER -->
        <div class="header">
          <div class="brand">BANGKOK <span>H</span>ORIZON</div>
          <div class="doc-title">
            <h1 class="doc-name">Booking Receipt</h1>
            <div class="doc-meta">#BOOK-${d.roomNumber}-${dateStr.replace(/\s/g,'')}</div>
            <div class="doc-meta">Date: ${dateStr} ${timeStr}</div>
          </div>
        </div>

        <!-- INFO GRID -->
        <div class="grid">
          <!-- CUSTOMER -->
          <div class="col">
            <div class="box-title">Billed To (ผู้เช่า)</div>
            <div class="box">
              <div class="info-row">
                <span class="info-label">Name</span>
                <span class="info-val">${d.customerName}</span>
              </div>
              <div class="info-row">
                <span class="info-label">Phone</span>
                <span class="info-val">${d.phone}</span>
              </div>
            </div>
          </div>

          <!-- PROPERTY -->
          <div class="col">
            <div class="box-title">Property Details (ห้องชุด)</div>
            <div class="box highlight-box">
              <div class="info-row">
                <span class="info-label">Room No.</span>
                <span class="info-val">${d.roomNumber}</span>
              </div>
              <div class="info-row">
                <span class="info-label">Type</span>
                <span class="info-val">${d.type}</span>
              </div>
              <div class="info-row">
                <span class="info-label">Floor / Size</span>
                <span class="info-val">${d.floor} / ${d.size} sqm</span>
              </div>
            </div>
          </div>
        </div>

        <!-- PAYMENT -->
        <div class="box-title">Payment Details (ยอดชำระ)</div>
        <div style="border:1px solid #e2e8f0; border-radius:12px; overflow:hidden;">
            <table style="width:100%; border-collapse:collapse;">
                <tr style="background:#f8fafc; border-bottom:1px solid #e2e8f0;">
                    <th style="text-align:left; padding:20px; font-size:14px; color:#64748b; text-transform:uppercase;">Description</th>
                    <th style="text-align:right; padding:20px; font-size:14px; color:#64748b; text-transform:uppercase;">Amount (THB)</th>
                </tr>
                <tr>
                    <td style="padding:25px; font-weight:600; color:#334155; font-size:16px;">
                        Booking Fee (เงินมัดจำจอง)
                        <div style="font-size:14px; font-weight:400; color:#94a3b8; margin-top:6px;">
                            Reservation for Room ${d.roomNumber}
                        </div>
                    </td>
                    <td style="text-align:right; padding:25px; font-weight:700; font-size:20px;">
                        ${Number(d.bookingFee).toLocaleString()}
                    </td>
                </tr>
            </table>
        </div>

        <div class="total-section">
            <div class="total-label">Total Amount Paid</div>
            <div class="total-amount">${Number(d.bookingFee).toLocaleString()} <span style="font-size:20px; font-weight:500; color:#94a3b8;">THB</span></div>
        </div>

        <!-- TERMS -->
        <div class="terms">
            <strong>Terms & Conditions:</strong><br>
            1. This booking fee is non-refundable (เงินจองไม่สามารถคืนได้).<br>
            2. The tenant must sign the rental contract within 7 days (ผู้เช่าต้องทำสัญญาภายใน 7 วัน).
        </div>

        <!-- SIGNATURES -->
        <div class="signature-area">
            <div class="sig-block">
                <div class="sig-line"></div>
                <div class="sig-name">${d.customerName}</div>
                <div class="sig-role">Customer (ผู้เช่า)</div>
            </div>
            <div class="sig-block">
                <div class="sig-line"></div>
                <div class="sig-name">Authorized Agent</div>
                <div class="sig-role">Bangkok Horizon (ผู้ดูแล)</div>
            </div>
        </div>

      </div>
    </body>
    </html>
  `;
}
