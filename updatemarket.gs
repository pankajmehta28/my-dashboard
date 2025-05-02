function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page');
}

function getBlankMarketsWithOptions() {
  const sheetId = '1LeoUoMWhwqGCQnlDN6qAqOQjS-nYhwvZhsLP1reBtPw'; // Market Mapping Sheet ID
  const mappingSheetName = 'Sheet1'; // Market mapping sheet
  const masterSheetName = 'Market Master List'; // Market master list sheet
  
  const ss = SpreadsheetApp.openById(sheetId);

  // Fetch blank markets
  const mappingSheet = ss.getSheetByName(mappingSheetName);
  const mappingData = mappingSheet.getDataRange().getValues();
  const headers = mappingData[0].map(h => (h || "").toString().trim());
  const billToCityIndex = headers.findIndex(h => h.toLowerCase() === 'bill to city');
  const marketIndex = headers.findIndex(h => h.toLowerCase() === 'market');

  if (billToCityIndex === -1 || marketIndex === -1) {
    throw new Error("❌ 'Bill To City' or 'Market' column not found. Check your sheet headers!");
  }

  const blankMarkets = [];
  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];
    const billToCity = (row[billToCityIndex] || "").toString().trim();
    const market = (row[marketIndex] || "").toString().trim();
    if (billToCity && market === "") {
      blankMarkets.push({ row: i + 1, billToCity });
    }
  }
   // Sort the cities alphabetically
        blankMarkets.sort((a, b) => a.billToCity.localeCompare(b.billToCity));

  // Fetch market options from the master list
  const masterSheet = ss.getSheetByName(masterSheetName);
  const masterData = masterSheet.getDataRange().getValues();
  const marketOptions = masterData.map(row => row[0]).filter(option => option); // First column, remove blanks
  
  return { blankMarkets, marketOptions };
}

function updateMarkets(updatedMarkets) {
  const sheetId = '1LeoUoMWhwqGCQnlDN6qAqOQjS-nYhwvZhsLP1reBtPw';
  const sheetName = 'Sheet1';
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || "").toString().trim());
  const marketIndex = headers.findIndex(h => h.toLowerCase() === 'market');

  updatedMarkets.forEach(entry => {
    const rowNumber = entry.row;
    const newMarket = entry.market;
    sheet.getRange(rowNumber, marketIndex + 1).setValue(newMarket);
  });

  return '✅ Markets updated successfully!';
}

// HTML content is embedded within the Apps Script as a string for ease of deployment
function doGet() {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
          }
          th, td {
            padding: 8px;
            border: 1px solid #ccc;
            text-align: left;
          }
          th {
            background-color: #28a745;
            color: white; /* makes the text readable on green */
          }
          select {
            width: 100%;
            padding: 4px;
          }
          button {
            margin-top: 10px;
            padding: 8px 12px;
            border: none;
            cursor: pointer;
          }
          button:hover {
            opacity: 0.8; /* Add hover effect */
          }
          .update-button {
            background-color: #007BFF;
            color: white;
          }
          .refresh-button {
            background-color: orange;
            color: white;
          }
          .home-button {
            background-color: #FF5733; /* Red for Home button */
            color: white;
          }
          h2 {
            color: black;
            text-align: center;
          }
        </style>
      </head>
      <body>
        <!-- Home Button -->
        <h2>Update Market for Bill To Cities</h2>

        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 10px;">
          <!-- Left-aligned buttons -->
          <div>
            <button onclick="submitUpdates()" class="update-button px-4 py-2 font-bold rounded shadow hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-400">
              Update Markets
            </button>
            <button onclick="fetchBlankMarkets()" class="refresh-button px-4 py-2 font-bold rounded shadow hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-400">
              Refresh
            </button>
          </div>

          <!-- Right-aligned Home button -->
          <div>
            <button onclick="goBack()" class="home-button px-4 py-2 font-bold rounded shadow hover:bg-red-600 focus:outline-none focus:ring-2 focus:ring-red-400">
              ← Home
            </button>
          </div>
        </div>

        <div id="marketTable"></div>

        <script>
          function goBack() {
            window.location.href = "protected.html";
          }

          // Fetch the blank markets and populate the table
          function fetchBlankMarkets() {
            google.script.run.withSuccessHandler(function({ blankMarkets, marketOptions }) {
              if (blankMarkets.length === 0) {
                document.getElementById('marketTable').innerHTML = '<p>No blank market entries found!</p>';
                return;
              }

              let html = '<table>';
              html += '<tr><th>Bill To City</th><th>New Market</th></tr>';
              blankMarkets.forEach(row => {
                html += \`<tr>
                  <td>\${row.billToCity}</td>
                  <td>
                    <select id="market-\${row.row}">
                      <option value="">-- Select Market --</option>
                      \${marketOptions.map(option => \`<option value="\${option}">\${option}</option>\`).join('')}
                    </select>
                  </td>
                </tr>\`;
              });
              html += '</table>';

              document.getElementById('marketTable').innerHTML = html;
            }).getBlankMarketsWithOptions();
          }

          // Submit updates to the backend
          function submitUpdates() {
            const selects = document.querySelectorAll('select[id^="market-"]');
            const updates = [];

            selects.forEach(select => {
              const rowId = select.id.split('-')[1];
              const marketValue = select.value.trim();
              if (marketValue) {
                updates.push({ row: parseInt(rowId), market: marketValue });
              }
            });

            if (updates.length === 0) {
              alert('Please select at least one market value to update.');
              return;
            }

            google.script.run.withSuccessHandler(function(response) {
              alert(response);
              fetchBlankMarkets(); // Refresh the table after updating
            }).updateMarkets(updates);
          }

          // Fetch data when the page loads
          fetchBlankMarkets();
        </script>
      </body>
    </html>
  `);
}
