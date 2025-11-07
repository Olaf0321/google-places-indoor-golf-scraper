/**
 * IndoorGolf Collector for Kanto (GAS)
 *
 * Features:
 * - onOpen menu: Set API Key, Start Collection, Save CSV, Clear Triggers
 * - Stores API key in Script Properties (PropertiesService)
 * - Uses NEW Google Places API (v1) - requires Places API (New) to be enabled
 * - Two-phase collection to avoid 6-minute limit:
 *   PHASE "search": Text Search across centers & keywords -> append deduped minimal rows to Results sheet
 *   PHASE "details": For rows missing details, call Place Details in batches and update rows
 * - Batch sizes and pause intervals configurable
 * - Saves progress in Script Properties and uses time-based triggers to continue
 * - CSV export of Results to Drive
 *
 * NOTE: Replace nothing in code for API key; use the menu "IndoorGolf > Set API Key" to store your key securely.
 * IMPORTANT: You must enable "Places API (New)" in Google Cloud Console, not the legacy Places API.
 */

/* =========================
   CONFIGURATION
   ========================= */
const CONFIG = {
  SHEET_NAME: 'IndoorGolf',
  RESULTS_SHEET: 'Results',
  // Search keywords to use (Japanese + English fallback)
  KEYWORDS: [
    'インドアゴルフ', 'ゴルフ練習場', 'ゴルフスクール', 'ゴルフレンジ', 'ゴルフ打ちっ放し',
    'indoor golf', 'golf practice', 'golf school', 'golf range', 'golf center',
    'ゴルフ', 'golf'
  ],
  // Kanto 7 prefectures cover centers (lat,lng). Add more for better coverage if required.
  // These are representative major-city centers; you may expand with a grid later.
  CENTERS: [
    { name: '東京', lat: 35.6809591, lng: 139.7673068 },
    { name: '神奈川', lat: 35.4437078, lng: 139.6380256 }, // 横浜
    // // Additional centers for better coverage
    // { name: 'Shibuya', lat: 35.658034, lng: 139.701636 },
    // { name: 'Shinjuku', lat: 35.689487, lng: 139.691706 },
    // { name: 'Ikebukuro', lat: 35.729503, lng: 139.715902 },
    // { name: 'Ueno', lat: 35.714759, lng: 139.777254 },
    // { name: 'Odaiba', lat: 35.630001, lng: 139.780000 },
    // { name: 'Machida', lat: 35.540000, lng: 139.450000 },
    // { name: 'Tachikawa', lat: 35.710000, lng: 139.410000 },
    // { name: 'Hachioji', lat: 35.660000, lng: 139.320000 }
  ],
  SEARCH_RADIUS_METERS: 25000, // Increased to 25km for better coverage
  BATCH_SIZE_SEARCH: 100,      // number of search results processed per batch before re-trigger
  BATCH_SIZE_DETAILS: 80,      // number of Place Details calls per batch
  NEXT_PAGE_WAIT_MS: 2500,     // wait before using next_page_token (approx 2s)
  DETAILS_SLEEP_MS: 250,       // sleep between details calls to avoid strict rate limits
  TRIGGER_RETRY_MINUTES: 1,    // next trigger will run in ~1 minute (minimum)
  PROPS_PREFIX: 'IndoorGolf_'
};

/* =========================
   UTILS: Properties / Triggers
   ========================= */
function getScriptProps() {
  return PropertiesService.getScriptProperties();
}
function getUserProps() {
  return PropertiesService.getUserProperties();
}
function getDocProps() {
  return PropertiesService.getDocumentProperties();
}
function setProp(key, value) {
  const p = getScriptProps();
  p.setProperty(CONFIG.PROPS_PREFIX + key, typeof value === 'string' ? value : JSON.stringify(value));
}
function getProp(key) {
  const raw = getScriptProps().getProperty(CONFIG.PROPS_PREFIX + key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return raw; }
}
function deleteProp(key) {
  getScriptProps().deleteProperty(CONFIG.PROPS_PREFIX + key);
}
function clearAllIndoorGolfProps() {
  // Preserve API key robustly: capture it, clear all prefixed keys EXCEPT API key, then restore
  const p = getScriptProps();
  const up = getUserProps();
  const dp = getDocProps ? getDocProps() : PropertiesService.getDocumentProperties();
  const apiKeyName = CONFIG.PROPS_PREFIX + 'API_KEY';
  
  // Capture API key from all possible locations
  const existingKey = p.getProperty(apiKeyName) || up.getProperty(apiKeyName) || dp.getProperty(apiKeyName) || '';
  
  // Clear all prefixed properties EXCEPT the API key
  const allProps = p.getProperties();
  Object.keys(allProps).forEach(function(keyName) {
    if (keyName.startsWith(CONFIG.PROPS_PREFIX) && keyName !== apiKeyName) {
      p.deleteProperty(keyName);
    }
  });
  
  // Ensure API key is stored in all locations for redundancy
  if (existingKey) {
    try { p.setProperty(apiKeyName, existingKey); } catch (e) { Logger.log('Failed to store API key in Script Properties: ' + e); }
    try { up.setProperty(apiKeyName, existingKey); } catch (e) { Logger.log('Failed to store API key in User Properties: ' + e); }
    try { dp.setProperty(apiKeyName, existingKey); } catch (e) { Logger.log('Failed to store API key in Document Properties: ' + e); }
  }
}

// Clear only run-time progress/state, never the API key
function clearProgressOnly() {
  // Only clear specific progress properties, never the API key
  const propertiesToClear = ['STATE', 'BATCH_STATE', 'SEARCH_PROGRESS', 'DETAILS_PROGRESS'];
  propertiesToClear.forEach(function(propName) {
    try { deleteProp(propName); } catch (e) { Logger.log('Failed to clear property ' + propName + ': ' + e); }
  });
}
/* =========================
  TYPE
  ========================== */

function getPlaceTypeCategory(types) {
  if (!types || types.length === 0) return '';

  // --- アウトドア（屋外系） ---
  const outdoorKeywords = [
    'golf_course',
    'park',
    'campground',
    'stadium',
    'zoo',
    'amusement_park',
    'rv_park',
    'beach',
    'natural_feature',
    'athletic_field',
    'sports_complex',
    'sports_activity_location',
    'marina',
    'hiking_area',
    'ski_resort',
    'surf_spot',
    'outdoor_recreation',
    'climbing_area',
    'cycling_route',
    'sports_coaching',
  ];

  // --- インドア（屋内系） ---
  const indoorKeywords = [
    'gym',
    'fitness_center',
    'yoga_studio',
    'spa',
    'bowling_alley',
    'movie_theater',
    'shopping_mall',
    'museum',
    'sports_club',
    'health',
    'massage',
    'store',
    'sporting_goods_store',
    'indoor_play_area',
    'dance_studio',
    'dojo',
    'indoor_arena',
    'pool',
  ];

  // --- 判定 ---
  const hasOutdoor = types.some(t => outdoorKeywords.includes(t));
  const hasIndoor = types.some(t => indoorKeywords.includes(t));

  if (hasOutdoor && !hasIndoor) return 'アウトドア';
  if (hasIndoor && !hasOutdoor) return 'インドア';
  if (hasOutdoor && hasIndoor) return '複合（アウトドア＋インドア）';

  return 'その他';
}

/* =========================
   MENU
   ========================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Attempt to repair API key if script properties were cleared
  try { repairApiKeyIfMissing(); } catch (e) {}
  ui.createMenu('IndoorGolf')
    .addItem('Set API Key', 'uiSetApiKey')
    .addItem('関東収集（実行）', 'startCollection')
    .addItem('CSV保存', 'saveAsCSV')
    // .addSeparator()
    // .addItem('Diagnose API Key', 'diagnoseApiKey')
    // .addItem('Fix API Issues', 'fixApiIssues')
    // .addItem('Debug Test Ultra Minimal API', 'debugTestUltraMinimalAPI')
    // .addItem('Debug Test Basic API', 'debugTestBasicAPI')
    // .addItem('Debug Test Complete Flow', 'debugTestCompleteFlow')
    // .addItem('Debug Test Minimal API', 'debugTestMinimalAPI')
    // .addItem('Debug Test Simple API', 'debugTestSimpleAPI')
    // .addItem('Debug Test Search', 'debugTestSearch')
    // .addItem('Debug Test Details', 'debugTestDetails')
    // .addItem('Debug Test Data Storage', 'debugTestDataStorage')
    // .addItem('Debug Check Sheet Status', 'debugCheckSheetStatus')
    // .addItem('Debug Force Add Test Data', 'debugForceAddTestData')
    // .addItem('Clear Triggers & Progress', 'clearTriggersAndProgress')
    .addToUi();
}

/* =========================
   API KEY: store/retrieve
   ========================= */
function uiSetApiKey() {
  const ui = SpreadsheetApp.getUi();
  
  // Show instructions first
  const instructions = 'To get a Google Places API key:\n\n' +
    '1. Go to Google Cloud Console (console.cloud.google.com)\n' +
    '2. Create a new project or select existing one\n' +
    '3. Enable "Places API (New)" - NOT the old Places API\n' +
    '4. Go to "Credentials" and create an API key\n' +
    '5. Restrict the key to "Places API (New)"\n' +
    '6. Enable billing for the project\n' +
    '7. Copy the API key and paste it below\n\n' +
    'IMPORTANT: You need "Places API (New)" enabled, not the old Places API!';
  
  ui.alert('API Key Setup Instructions', instructions, ui.ButtonSet.OK);
  
  const resp = ui.prompt('Set Google Places API Key', 'Paste your Places API (New) key here:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const key = resp.getResponseText().trim();
    if (!key) { 
      ui.alert('No key entered.'); 
      return; 
    }
    
    // Validate key format
    if (key.length < 20) {
      ui.alert('API key seems too short. Please check and try again.');
      return;
    }
    
    // Persist to all stores (prefer User Properties for stability across triggers)
    try { getUserProps().setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', key); } catch (e) {}
    try { getDocProps().setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', key); } catch (e) {}
    setProp('API_KEY', key);
    
    // Test the key immediately
    ui.alert('API key saved. Testing...');
    
    try {
      const testConfig = buildSimpleSearchUrl(key, 'test', 35.6762, 139.6503);
      const resp = UrlFetchApp.fetch(testConfig.url, {
        method: testConfig.method,
        payload: testConfig.body,
        headers: testConfig.headers,
        muteHttpExceptions: true
      });
      
      if (resp.getResponseCode() === 200) {
        ui.alert('✅ SUCCESS! API key is working correctly.');
      } else if (resp.getResponseCode() === 403) {
        ui.alert('❌ API key saved but has permission issues (403).\n\n' +
                'Please check:\n' +
                '1. Places API (New) is enabled\n' +
                '2. Billing is enabled\n' +
                '3. API key has correct permissions\n\n' +
                'Use "Diagnose API Key" for more details.');
      } else {
        ui.alert('⚠️ API key saved but test failed with code: ' + resp.getResponseCode() + '\n\n' +
                'Use "Diagnose API Key" for more details.');
      }
    } catch (e) {
      ui.alert('❌ API key saved but test failed: ' + e.toString() + '\n\n' +
              'Use "Diagnose API Key" for more details.');
    }
  } else {
    ui.alert('Cancelled.');
  }
}
function getApiKey() {
  // Backward-compatible helper: return key if available; do not throw
  const k = tryGetApiKey();
  return k; // may be null; callers should handle
}
// Non-throwing fetch used by background triggers to avoid noisy alerts
function tryGetApiKey() {
  try {
    // Prefer User Properties → Script Properties → Document Properties
    const up = getUserProps();
    let k = up.getProperty(CONFIG.PROPS_PREFIX + 'API_KEY');
    if (k) { setProp('API_KEY', k); return k; }
    k = getScriptProps().getProperty(CONFIG.PROPS_PREFIX + 'API_KEY');
    if (k) { try { up.setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', k); } catch (e) {} return k; }
    const dp = getDocProps();
    k = dp.getProperty(CONFIG.PROPS_PREFIX + 'API_KEY');
    if (k) { setProp('API_KEY', k); try { up.setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', k); } catch (e) {} return k; }
    return null;
  } catch (e) {
    return null;
  }
}

// Self-heal: if the Script property is missing but a backup exists, restore it
function repairApiKeyIfMissing() {
  const sp = getScriptProps();
  const name = CONFIG.PROPS_PREFIX + 'API_KEY';
  if (sp.getProperty(name)) return;
  const up = getUserProps();
  const dp = getDocProps();
  const backup = up.getProperty(name) || dp.getProperty(name);
  if (backup) sp.setProperty(name, backup);
}

// Ensure a key exists; prompts the user if missing
function ensureApiKeyInteractive() {
  let key = tryGetApiKey();
  if (key) return key;
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Google Places API Key required', 'Paste your API key now:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  const entered = (resp.getResponseText() || '').trim();
  if (!entered) return null;
  setProp('API_KEY', entered);
  try { getUserProps().setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', entered); } catch (e) {}
  try { getDocProps().setProperty(CONFIG.PROPS_PREFIX + 'API_KEY', entered); } catch (e) {}
  return entered;
}

// Simple diagnostic to verify where the API key is stored
function diagnoseApiKey() {
  const name = CONFIG.PROPS_PREFIX + 'API_KEY';
  const sp = getScriptProps().getProperty(name) ? 'yes' : 'no';
  const up = getUserProps().getProperty(name) ? 'yes' : 'no';
  const dp = getDocProps().getProperty(name) ? 'yes' : 'no';
  
  let message = 'API Key presence:\n';
  message += 'Script Properties: ' + sp + '\n';
  message += 'User Properties: ' + up + '\n';
  message += 'Document Properties: ' + dp + '\n\n';
  
  if (sp === 'no' && up === 'no' && dp === 'no') {
    message += '❌ NO API KEY FOUND!\n';
    message += 'Please use "Set API Key" to add your Google Places API key.';
  } else {
    const apiKey = tryGetApiKey();
    if (apiKey) {
      message += '✅ API Key found: ' + apiKey.substring(0, 10) + '...\n';
      message += 'Key length: ' + apiKey.length + ' characters\n\n';
      
      // Test if the key works
      message += 'Testing API key...\n';
      try {
        const testConfig = buildSimpleSearchUrl(apiKey, 'test', 35.6762, 139.6503);
        const resp = UrlFetchApp.fetch(testConfig.url, {
          method: testConfig.method,
          payload: testConfig.body,
          headers: testConfig.headers,
          muteHttpExceptions: true
        });
        
        if (resp.getResponseCode() === 200) {
          message += '✅ API Key is working!';
        } else if (resp.getResponseCode() === 403) {
          message += '❌ API Key has permission issues (403)\n';
          message += 'Check if Places API (New) is enabled and billing is set up.';
        } else {
          message += '⚠️ API Key test failed with code: ' + resp.getResponseCode();
        }
      } catch (e) {
        message += '❌ API Key test error: ' + e.toString();
      }
    } else {
      message += '❌ API Key could not be retrieved!';
    }
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// Function to help fix API issues
function fixApiIssues() {
  const ui = SpreadsheetApp.getUi();
  
  const instructions = 'Common API Issues and Solutions:\n\n' +
    '❌ 403 ERROR - Permission Denied:\n' +
    '1. Go to Google Cloud Console\n' +
    '2. Select your project\n' +
    '3. Go to "APIs & Services" > "Library"\n' +
    '4. Search for "Places API (New)"\n' +
    '5. Click on it and press "Enable"\n' +
    '6. Go to "APIs & Services" > "Credentials"\n' +
    '7. Click on your API key\n' +
    '8. Under "API restrictions", select "Places API (New)"\n' +
    '9. Go to "Billing" and make sure billing is enabled\n\n' +
    '❌ 400 ERROR - Bad Request:\n' +
    '1. Check if the API key is valid\n' +
    '2. Make sure you\'re using Places API (New), not the old one\n\n' +
    '❌ 429 ERROR - Rate Limit:\n' +
    '1. Wait a few minutes and try again\n' +
    '2. Check your quota limits in Google Cloud Console\n\n' +
    'After fixing, use "Diagnose API Key" to test.';
  
  ui.alert('API Issues Fix Guide', instructions, ui.ButtonSet.OK);
}

// Test function to verify data storage is working
function debugTestDataStorage() {
  try {
    const sheet = ensureSheets();
    const lastRowBefore = sheet.getLastRow();
    
    Logger.log('debugTestDataStorage - Before: %s rows', lastRowBefore);
    
    // Create a test row
    const testRowObj = {
      店舗ID: 'TEST_PLACE_123',
      店舗名: 'Test Golf Center',
      住所: '123 Test Street, 東京',
      // vicinity: '123 Test Street, 東京',
      // lat: '35.6762',
      // lng: '139.6503',
      評価: '4.5',
      口コミ数: '100',
      電話番号: '03-1234-5678',
      WebサイトURL: 'https://test.com',
      // opening_hours_weekday_text: 'Mon-Fri: 9AM-9PM',
      // open_now: 'true',
      // types: 'golf_course,sports_complex',
      // 種別: getPlaceTypeCategory(testPlace.types),
      種別: 'アウトドア',
      // url: 'https://maps.google.com/test',
      営業ステータス: '運用',
      検索に使った地域名: '東京',
      検索に使ったキーワード: 'test',
      // collected_at: new Date().toISOString(),
      // details_fetched: 'no'
    };
    
    // Try to add the test row
    if (testRowObj['検索に使った地域名'] ==  '東京' || testRowObj['検索に使った地域名'] ==  '神奈川') {
      appendRowInOrder(sheet, testRowObj);
    }
    
    const lastRowAfter = sheet.getLastRow();
    Logger.log('debugTestDataStorage - After: %s rows', lastRowAfter);
    
    if (lastRowAfter > lastRowBefore) {
      // Verify the data was actually written
      const addedRow = sheet.getRange(lastRowAfter, 1, 1, sheet.getLastColumn()).getValues()[0];
      const hasData = addedRow.some(cell => cell && cell.toString().trim() !== '');
      
      if (hasData) {
        SpreadsheetApp.getUi().alert('SUCCESS: Test row added! Check the Results sheet for "Test Golf Center"');
      } else {
        SpreadsheetApp.getUi().alert('WARNING: Row was added but appears empty. Check the logs for details.');
      }
    } else {
      SpreadsheetApp.getUi().alert('FAILED: No row was added. Check the logs for errors.');
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Test failed with error: ' + e.toString());
    Logger.log('debugTestDataStorage ERROR: %s', e.toString());
  }
}

// New function to check current sheet status
function debugCheckSheetStatus() {
  try {
    const sheet = ensureSheets();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    Logger.log('Sheet Status - Rows: %s, Columns: %s', lastRow, lastCol);
    
    if (lastRow > 1) {
      const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
      Logger.log('Headers: %s', headers.join(' | '));
      
      // Show first few data rows
      const dataRows = sheet.getRange(2,1,Math.min(3, lastRow-1), lastCol).getValues();
      dataRows.forEach((row, index) => {
        Logger.log('Row %s: %s', index + 2, row.join(' | '));
      });
    }
    
    SpreadsheetApp.getUi().alert('Sheet Status:\nRows: ' + lastRow + '\nColumns: ' + lastCol + '\nCheck Logger for details');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error checking sheet status: ' + e.toString());
    Logger.log('debugCheckSheetStatus ERROR: %s', e.toString());
  }
}

// Test complete data flow from API to sheet
function debugTestCompleteFlow() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    Logger.log('Testing complete data flow from API to sheet...');
    
    // Test with minimal API call
    const base = 'https://places.googleapis.com/v1/places:searchText';
    const params = [];
    params.push('key=' + encodeURIComponent(apiKey));
    
    const requestBody = {
      textQuery: 'golf',
      locationBias: {
        circle: {
          center: {
            latitude: 35.6762,
            longitude: 139.6503
          },
          radius: 10000
        }
      }
    };
    
    Logger.log('Making API call...');
    const resp = UrlFetchApp.fetch(base + '?' + params.join('&'), {
      method: 'POST',
      body: JSON.stringify(requestBody),
      headers: {
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const js = JSON.parse(resp.getContentText());
    Logger.log('API Response Code: %s', resp.getResponseCode());
    Logger.log('API Response: %s', JSON.stringify(js, null, 2));
    
    if (resp.getResponseCode() === 200 && js.places && js.places.length > 0) {
      Logger.log('API call successful, found %s places', js.places.length);
      
      // Test sheet writing
      const sheet = ensureSheets();
      const lastRowBefore = sheet.getLastRow();
      Logger.log('Sheet has %s rows before adding data', lastRowBefore);
      
      // Try to add the first place to the sheet
      const testPlace = js.places[0];
      const rowObj = {
        店舗ID: testPlace.id || 'TEST_' + Date.now(),
        店舗名: testPlace.displayName['text'] || 'Test Place',
        住所: testPlace.formattedAddress || 'Test Address',
        種別: getPlaceTypeCategory(testPlace.types),
        電話番号: testPlace.nationalPhoneNumber || '',
        WebサイトURL: testPlace.websiteUri || '',
        // vicinity: testPlace.formattedAddress || 'Test Address',
        // lat: (testPlace.location && testPlace.location.latitude) || '35.6762',
        // lng: (testPlace.location && testPlace.location.longitude) || '139.6503',
        評価: testPlace.rating || '0',
        口コミ数: testPlace.userRatingCount || '0',
        // opening_hours_weekday_text: '',
        // open_now: (testPlace.currentOpeningHours && typeof testPlace.currentOpeningHours.openNow !== 'undefined') ? String(testPlace.currentOpeningHours.openNow) : '',
        // types: (testPlace.types || []).join(','),

        // url: testPlace.googleMapsUri || '',
        営業ステータス: testPlace.businessStatus || '',
        検索に使った地域名: '東京',
        検索に使ったキーワード: 'golf',
        // collected_at: new Date().toISOString(),
        // details_fetched: 'no'
      };
      
      Logger.log('About to add test place to sheet: %s', rowObj.店舗名);
      
      try {
        appendRowInOrder(sheet, rowObj);
        
        const lastRowAfter = sheet.getLastRow();
        Logger.log('Sheet has %s rows after adding data', lastRowAfter);
        
        if (lastRowAfter > lastRowBefore) {
          SpreadsheetApp.getUi().alert('✅ SUCCESS! Complete data flow works!\n\n' +
            'API found: ' + js.places.length + ' places\n' +
            'Added to sheet: 1 place\n' +
            'Check the Results sheet for the test data.');
        } else {
          SpreadsheetApp.getUi().alert('❌ FAILED: No data was added to the sheet.\n\n' +
            'API found: ' + js.places.length + ' places\n' +
            'But sheet writing failed.\n' +
            'Check the logs for details.');
        }
      } catch (e) {
        SpreadsheetApp.getUi().alert('❌ FAILED: Error adding data to sheet: ' + e.toString());
        Logger.log('Sheet writing error: %s', e.toString());
      }
      
    } else {
      SpreadsheetApp.getUi().alert('❌ API call failed with code: ' + resp.getResponseCode() + '\n\nResponse: ' + resp.getContentText());
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Complete flow test error: ' + e.toString());
    Logger.log('debugTestCompleteFlow ERROR: %s', e.toString());
  }
}

// Test API with the most basic possible request
function debugTestBasicAPI() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    Logger.log('Testing API with most basic request...');
    
    // Use the most basic possible request format
    const base = 'https://places.googleapis.com/v1/places:searchText';
    const params = [];
    params.push('key=' + encodeURIComponent(apiKey));
    
    // Minimal request body - only required fields
    const requestBody = {
      textQuery: 'golf',
      locationBias: {
        circle: {
          center: {
            latitude: 35.6762,
            longitude: 139.6503
          },
          radius: 10000
        }
      }
    };
    
    Logger.log('Basic Request Body: %s', JSON.stringify(requestBody, null, 2));
    Logger.log('API URL: %s', base + '?' + params.join('&'));
    
    const resp = UrlFetchApp.fetch(base + '?' + params.join('&'), {
      method: 'POST',
      body: JSON.stringify(requestBody),
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.location'
      },
      muteHttpExceptions: true
    });
    
    const js = JSON.parse(resp.getContentText());
    
    Logger.log('Basic API Test - Response Code: %s', resp.getResponseCode());
    Logger.log('Basic API Test - Response: %s', JSON.stringify(js, null, 2));
    
    if (resp.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert('✅ SUCCESS! Basic API call works!\n\n' +
        'Found: ' + (js.places ? js.places.length : 0) + ' places\n' +
        'Check Logger for full response details.');
    } else {
      let errorMsg = '❌ Basic API call failed with code: ' + resp.getResponseCode() + '\n\n';
      
      if (resp.getResponseCode() === 403) {
        errorMsg += '403 ERROR - Permission Denied\n\n';
        errorMsg += 'Your API key is configured correctly, but there might be:\n';
        errorMsg += '1. Billing not enabled for the project\n';
        errorMsg += '2. API key restrictions too strict\n';
        errorMsg += '3. Quota limits exceeded\n\n';
        errorMsg += 'Please check your Google Cloud Console billing and quotas.';
      } else if (resp.getResponseCode() === 400) {
        errorMsg += '400 ERROR - Bad Request\n\n';
        errorMsg += 'The request format might be incorrect.';
      }
      
      errorMsg += '\n\nFull Response: ' + resp.getContentText();
      SpreadsheetApp.getUi().alert(errorMsg);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Basic API Test Error: ' + e.toString());
    Logger.log('debugTestBasicAPI ERROR: %s', e.toString());
  }
}

// Test API with absolutely minimal request (no field mask)
function debugTestUltraMinimalAPI() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    Logger.log('Testing API with ultra minimal request (no field mask)...');
    
    // Use the absolute minimal request format
    const base = 'https://places.googleapis.com/v1/places:searchText';
    const params = [];
    params.push('key=' + encodeURIComponent(apiKey));
    
    // Ultra minimal request body
    const requestBody = {
      textQuery: 'golf',
      locationBias: {
        circle: {
          center: {
            latitude: 35.6762,
            longitude: 139.6503
          },
          radius: 10000
        }
      }
    };
    
    Logger.log('Ultra Minimal Request Body: %s', JSON.stringify(requestBody, null, 2));
    Logger.log('API URL: %s', base + '?' + params.join('&'));
    
    const resp = UrlFetchApp.fetch(base + '?' + params.join('&'), {
      method: 'POST',
      body: JSON.stringify(requestBody),
      headers: {
        'Content-Type': 'application/json'
        // No field mask at all
      },
      muteHttpExceptions: true
    });
    
    const js = JSON.parse(resp.getContentText());
    
    Logger.log('Ultra Minimal API Test - Response Code: %s', resp.getResponseCode());
    Logger.log('Ultra Minimal API Test - Response: %s', JSON.stringify(js, null, 2));
    
    if (resp.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert('✅ SUCCESS! Ultra minimal API call works!\n\n' +
        'Found: ' + (js.places ? js.places.length : 0) + ' places\n' +
        'This means the field mask was causing the 403 error.\n' +
        'Check Logger for full response details.');
    } else {
      let errorMsg = '❌ Ultra minimal API call failed with code: ' + resp.getResponseCode() + '\n\n';
      
      if (resp.getResponseCode() === 403) {
        errorMsg += '403 ERROR - Still getting permission denied\n\n';
        errorMsg += 'This suggests the issue is with:\n';
        errorMsg += '1. Billing not enabled for the project\n';
        errorMsg += '2. API key restrictions\n';
        errorMsg += '3. Project configuration\n\n';
        errorMsg += 'Please check your Google Cloud Console billing.';
      }
      
      errorMsg += '\n\nFull Response: ' + resp.getContentText();
      SpreadsheetApp.getUi().alert(errorMsg);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ultra Minimal API Test Error: ' + e.toString());
    Logger.log('debugTestUltraMinimalAPI ERROR: %s', e.toString());
  }
}

// Test API with minimal request
function debugTestMinimalAPI() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    Logger.log('Testing API with minimal request...');
    
    // Test with the most basic possible request
    const base = 'https://places.googleapis.com/v1/places:searchText';
    const params = [];
    params.push('key=' + encodeURIComponent(apiKey));
    
    const requestBody = {
      textQuery: 'golf',
      locationBias: {
        circle: {
          center: {
            latitude: 35.6762,
            longitude: 139.6503
          },
          radius: 10000
        }
      }
    };
    
    Logger.log('Minimal Request Body: %s', JSON.stringify(requestBody, null, 2));
    
    const resp = UrlFetchApp.fetch(base + '?' + params.join('&'), {
      method: 'POST',
      body: JSON.stringify(requestBody),
      headers: {
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const js = JSON.parse(resp.getContentText());
    
    Logger.log('Minimal API Test - Response Code: %s', resp.getResponseCode());
    Logger.log('Minimal API Test - Response: %s', JSON.stringify(js, null, 2));
    
    if (resp.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert('✅ SUCCESS! Minimal API call works! Found ' + (js.places ? js.places.length : 0) + ' places.');
    } else {
      SpreadsheetApp.getUi().alert('❌ Minimal API call failed with code: ' + resp.getResponseCode() + '\n\nResponse: ' + resp.getContentText());
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Minimal API Test Error: ' + e.toString());
    Logger.log('debugTestMinimalAPI ERROR: %s', e.toString());
  }
}

// Test API with very simple search
function debugTestSimpleAPI() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    Logger.log('Testing API with very simple search...');
    
    // Test with just "golf" in 東京
    const testConfig = buildSimpleSearchUrl(apiKey, 'golf', 35.6762, 139.6503);
    const resp = UrlFetchApp.fetch(testConfig.url, {
      method: testConfig.method,
      payload: testConfig.body,
      headers: testConfig.headers,
      muteHttpExceptions: true
    });
    
    const js = JSON.parse(resp.getContentText());
    
    Logger.log('Simple API Test - Response Code: %s', resp.getResponseCode());
    Logger.log('Simple API Test - Response: %s', JSON.stringify(js, null, 2));
    
    if (resp.getResponseCode() === 200 && js.places && js.places.length > 0) {
      SpreadsheetApp.getUi().alert('SUCCESS: API is working! Found ' + js.places.length + ' places. Check Logger for details.');
      
      // Try to add the first result to the sheet
      const sheet = ensureSheets();
      const testPlace = js.places[0];
      const rowObj = {
        店舗ID: testPlace.id || 'TEST_' + Date.now(),
        店舗名: testPlace.displayName['text'] || 'Test Place',
        住所: testPlace.formattedAddress || 'Test Address',
        種別: getPlaceTypeCategory(testPlace.types),
        電話番号: testPlace.nationalPhoneNumber || '',
        WebサイトURL: testPlace.websiteUri || '',
        評価: testPlace.rating || '0',
        口コミ数: testPlace.userRatingCount || '0',
        // vicinity: testPlace.formattedAddress || 'Test Address',
        // lat: (testPlace.location && testPlace.location.latitude) || '35.6762',
        // lng: (testPlace.location && testPlace.location.longitude) || '139.6503',
        // opening_hours_weekday_text: '',
        // open_now: (testPlace.currentOpeningHours && typeof testPlace.currentOpeningHours.openNow !== 'undefined') ? String(testPlace.currentOpeningHours.openNow) : '',
        // types: (testPlace.types || []).join(','),
        // url: testPlace.googleMapsUri || '',
        営業ステータス: testPlace.businessStatus || '',
        検索に使った地域名: '東京',
        検索に使ったキーワード: 'golf',
        // collected_at: new Date().toISOString(),
        // details_fetched: 'no'
      };
      
      appendRowInOrder(sheet, rowObj);
      SpreadsheetApp.getUi().alert('Added test result to sheet! Check the Results sheet.');
      
    } else {
      let errorMessage = 'API Test Failed - Response Code: ' + resp.getResponseCode() + '\n';
      
      if (resp.getResponseCode() === 403) {
        errorMessage += '\n403 ERROR - This usually means:\n';
        errorMessage += '1. API key is invalid or missing\n';
        errorMessage += '2. Places API (New) is not enabled\n';
        errorMessage += '3. API key doesn\'t have Places API permissions\n';
        errorMessage += '4. Billing is not enabled for the project\n\n';
        errorMessage += 'Please check:\n';
        errorMessage += '- Go to Google Cloud Console\n';
        errorMessage += '- Enable Places API (New)\n';
        errorMessage += '- Check API key permissions\n';
        errorMessage += '- Verify billing is enabled';
      } else if (resp.getResponseCode() === 400) {
        errorMessage += '\n400 ERROR - Bad Request\n';
        errorMessage += 'Check the request parameters in the logs';
      } else if (resp.getResponseCode() === 429) {
        errorMessage += '\n429 ERROR - Rate Limit Exceeded\n';
        errorMessage += 'Wait a few minutes and try again';
      }
      
      errorMessage += '\n\nFull Response: ' + resp.getContentText();
      SpreadsheetApp.getUi().alert(errorMessage);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('API Test Error: ' + e.toString());
    Logger.log('debugTestSimpleAPI ERROR: %s', e.toString());
  }
}

// Force add test data to verify sheet writing works
function debugForceAddTestData() {
  try {
    const sheet = ensureSheets();
    const lastRowBefore = sheet.getLastRow();
    
    Logger.log('Force adding test data - Before: %s rows', lastRowBefore);
    
    // Add multiple test rows to verify the sheet writing works
    const testData = [
      {
        店舗ID: 'TEST_PLACE_001',
        店舗名: 'Test Golf Center 1',
        住所: '123 Test Street, 東京',
        種別: 'アウトドア',
        電話番号: '03-1234-5678',
        WebサイトURL: 'https://test1.com',
        評価: '4.5',
        口コミ数: '100',
        // vicinity: '123 Test Street, 東京',
        // lat: '35.6762',
        // lng: '139.6503',
        // opening_hours_weekday_text: 'Mon-Fri: 9AM-9PM',
        // open_now: 'true',
        // types: 'golf_course,sports_complex',
        // url: 'https://maps.google.com/test1',
        営業ステータス: '運用',
        検索に使った地域名: '東京',
        検索に使ったキーワード: 'test',
        // collected_at: new Date().toISOString(),
        // details_fetched: 'no'
      },
      {
        店舗ID: 'TEST_PLACE_002',
        店舗名: 'Test Golf Center 2',
        住所: '456 Test Avenue, Yokohama',
        種別: 'アウトドア',
        電話番号: '045-1234-5678',
        WebサイトURL: 'https://test2.com',
        評価: '4.2',
        口コミ数: '85',
        // vicinity: '456 Test Avenue, Yokohama',
        lat: '35.4437',
        lng: '139.6380',
        // opening_hours_weekday_text: 'Mon-Sun: 8AM-10PM',
        // open_now: 'false',
        // types: 'golf_course,recreation',
        // url: 'https://maps.google.com/test2',
        営業ステータス: '運用',
        検索に使った地域名: 'Yokohama',
        検索に使ったキーワード: 'test',
        // collected_at: new Date().toISOString(),
        // details_fetched: 'no'
      }
    ];
    
    // Add each test row
    testData.forEach((rowObj, index) => {
      Logger.log('Adding test row %s: %s', index + 1, rowObj.店舗名);
      appendRowInOrder(sheet, rowObj);
    });
    
    const lastRowAfter = sheet.getLastRow();
    Logger.log('Force adding test data - After: %s rows', lastRowAfter);
    
    if (lastRowAfter > lastRowBefore) {
      SpreadsheetApp.getUi().alert('SUCCESS: Added ' + (lastRowAfter - lastRowBefore) + ' test rows! Check the Results sheet.');
    } else {
      SpreadsheetApp.getUi().alert('FAILED: No rows were added. Check the logs for errors.');
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Test failed with error: ' + e.toString());
    Logger.log('debugForceAddTestData ERROR: %s', e.toString());
  }
}

// Diagnostic function to test details fetching
function debugTestDetails() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    const sheet = ensureSheets();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      SpreadsheetApp.getUi().alert('No data in sheet. Run search first.');
      return;
    }
    
    const header = values[0];
    const rows = values.slice(1);
    
    // Find first row that needs details
    const colIndex = {};
    header.forEach((h, i) => colIndex[h] = i);
    
    let testRow = null;
    let testRowIndex = -1;
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const detailsFetched = row[colIndex['details_fetched']] || '';
      const pid = row[colIndex['店舗ID']];
      if ((detailsFetched === 'no' || detailsFetched === '') && pid) {
        testRow = row;
        testRowIndex = i + 2;
        break;
      }
    }
    
    if (!testRow) {
      SpreadsheetApp.getUi().alert('No rows found that need details fetching.');
      return;
    }
    
    const placeId = testRow[colIndex['店舗ID']];
    const placeName = testRow[colIndex['店舗名']] || 'Unknown';
    
    SpreadsheetApp.getUi().alert('Testing details fetch for:\nPlace: ' + placeName + '\nPlace ID: ' + placeId + '\nRow: ' + testRowIndex);
    
    // Test the details fetch
    const details = fetchPlaceDetails(apiKey, placeId);
    Logger.log('Debug Details - Fetched details: ' + JSON.stringify(details, null, 2));
    
    // Update the row
    const currentRow = sheet.getRange(testRowIndex, 1, 1, header.length).getValues()[0];
    
    if (colIndex['電話番号'] !== undefined) currentRow[colIndex['電話番号']] = details.電話番号 || currentRow[colIndex['電話番号']] || '';
    if (colIndex['WebサイトURL'] !== undefined) currentRow[colIndex['WebサイトURL']] = details.WebサイトURL || currentRow[colIndex['WebサイトURL']] || '';
    // if (colIndex['opening_hours_weekday_text'] !== undefined) currentRow[colIndex['opening_hours_weekday_text']] = details.opening_hours_weekday_text || currentRow[colIndex['opening_hours_weekday_text']] || '';
    // if (colIndex['open_now'] !== undefined) currentRow[colIndex['open_now']] = (typeof details.open_now !== 'undefined') ? String(details.open_now) : currentRow[colIndex['open_now']] || '';
    // if (colIndex['details_fetched'] !== undefined) currentRow[colIndex['details_fetched']] = 'yes';
    
    sheet.getRange(testRowIndex, 1, 1, header.length).setValues([currentRow]);
    
    SpreadsheetApp.getUi().alert('Details test completed. Check the sheet and Logger for results.');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Debug details test error: ' + e.toString());
    Logger.log('Debug details test error: ' + e.toString());
  }
}

// Diagnostic function to test API and sheet operations
function debugTestSearch() {
  try {
    const apiKey = tryGetApiKey();
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('No API key found. Please set it first.');
      return;
    }
    
    // Test with multiple search terms
    const testKeywords = ['インドアゴルフ', 'golf', 'ゴルフ練習場', 'golf practice'];
    const testCenter = CONFIG.CENTERS[0]; // 東京
    
    let totalFound = 0;
    let results = [];
    
    for (let testKeyword of testKeywords) {
      Logger.log('Testing keyword: %s', testKeyword);
      
      // Try with type restriction first
      const requestConfig = buildTextSearchUrl(apiKey, testKeyword, testCenter.lat, testCenter.lng, CONFIG.SEARCH_RADIUS_METERS, null, true);
      const resp = UrlFetchApp.fetch(requestConfig.url, {
        method: requestConfig.method,
        payload: requestConfig.body,
        headers: requestConfig.headers,
        muteHttpExceptions: true
      });
      
      const js = JSON.parse(resp.getContentText());
      
      Logger.log('Keyword: %s, Response Code: %s, Places found: %s', testKeyword, resp.getResponseCode(), js.places ? js.places.length : 0);
      Logger.log('Full Response: %s', JSON.stringify(js, null, 2));
      
      if (resp.getResponseCode() === 200 && js.places && js.places.length > 0) {
        totalFound += js.places.length;
        results = results.concat(js.places);
        Logger.log('Found %s places for keyword: %s', js.places.length, testKeyword);
      } else {
        // Try without type restriction
        Logger.log('No results with type restriction, trying broader search for: %s', testKeyword);
        const broaderConfig = buildTextSearchUrl(apiKey, testKeyword, testCenter.lat, testCenter.lng, CONFIG.SEARCH_RADIUS_METERS, null, false);
        const broaderResp = UrlFetchApp.fetch(broaderConfig.url, {
          method: broaderConfig.method,
          payload: broaderConfig.body,
          headers: broaderConfig.headers,
          muteHttpExceptions: true
        });
        
        const broaderJs = JSON.parse(broaderResp.getContentText());
        Logger.log('Broader search - Keyword: %s, Response Code: %s, Places found: %s', testKeyword, broaderResp.getResponseCode(), broaderJs.places ? broaderJs.places.length : 0);
        
        if (broaderResp.getResponseCode() === 200 && broaderJs.places && broaderJs.places.length > 0) {
          totalFound += broaderJs.places.length;
          results = results.concat(broaderJs.places);
          Logger.log('Found %s places with broader search for keyword: %s', broaderJs.places.length, testKeyword);
        }
      }
    }
    
    // Check sheet
    const sheet = ensureSheets();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    SpreadsheetApp.getUi().alert(
      'Debug Results:\n' +
      'Total places found: ' + totalFound + '\n' +
      'Sheet last row: ' + lastRow + '\n' +
      'Sheet last column: ' + lastCol + '\n' +
      'Check Logger for full API response details'
    );
    
    // If we have places, try to add one as a test
    if (results.length > 0) {
      const testPlace = results[0];
      const rowObj = {
        店舗ID: testPlace.id || '',
        店舗名: testPlace.displayName['text'] || '',
        住所: testPlace.formattedAddress || '',
        種別: getPlaceTypeCategory(testPlace.types),
        電話番号: testPlace.nationalPhoneNumber || '',
        WebサイトURL: testPlace.websiteUri || '',
        評価: testPlace.rating || '',
        口コミ数: testPlace.userRatingCount || '',
        // vicinity: testPlace.formattedAddress || '',
        // lat: (testPlace.location && testPlace.location.latitude) || '',
        // lng: (testPlace.location && testPlace.location.longitude) || '',
        // opening_hours_weekday_text: '',
        // open_now: (testPlace.currentOpeningHours && typeof testPlace.currentOpeningHours.openNow !== 'undefined') ? String(testPlace.currentOpeningHours.openNow) : '',
        // types: (testPlace.types || []).join(','),
        // url: testPlace.googleMapsUri || '',
        営業ステータス: testPlace.businessStatus || '',
        検索に使った地域名: testCenter.name,
        検索に使ったキーワード: 'test',
        // collected_at: new Date().toISOString(),
        // details_fetched: 'no'
      };
      
      appendRowInOrder(sheet, rowObj);
      SpreadsheetApp.getUi().alert('Added test row to sheet. Check the Results sheet.');
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Debug test error: ' + e.toString());
    Logger.log('Debug test error: ' + e.toString());
  }
}

/* =========================
   Setup: create sheets and headers
   ========================= */
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // create main sheet (for meta if needed)
  // if (!ss.getSheetByName(CONFIG.SHEET_NAME)) ss.insertSheet(CONFIG.SHEET_NAME);
  
  // create results sheet
  let results = ss.getSheetByName(CONFIG.RESULTS_SHEET);
  if (!results) {
    results = ss.insertSheet(CONFIG.RESULTS_SHEET);
    Logger.log('Created new Results sheet');
  }
  
  // prepare header row
  const headers = [
    '店舗ID', '店舗名', '住所','種別',
    '電話番号', 'WebサイトURL', '評価', '口コミ数',
    '営業ステータス', '検索に使った地域名', '検索に使ったキーワード'
  ];
  
  // Check if sheet has data
  const lastRow = results.getLastRow();
  const lastCol = results.getLastColumn();
  Logger.log('ensureSheets - Sheet has %s rows, %s columns', lastRow, lastCol);
  
  // Only set headers if sheet is empty or headers are missing
  if (lastRow === 0 || lastCol === 0) {
    Logger.log('ensureSheets - Setting headers on empty sheet');
    results.getRange(1,1,1,headers.length).setValues([headers]);
    SpreadsheetApp.flush(); // Force flush to ensure headers are written
  } else {
    // Check if headers match
    const currentHeaders = results.getRange(1,1,1,Math.min(lastCol, headers.length)).getValues()[0];
    const currentHeadersStr = currentHeaders.join('|');
    const expectedHeadersStr = headers.slice(0, currentHeaders.length).join('|');
    
    Logger.log('ensureSheets - Current headers: %s', currentHeadersStr);
    Logger.log('ensureSheets - Expected headers: %s', expectedHeadersStr);
    
    if (currentHeadersStr !== expectedHeadersStr) {
      Logger.log('ensureSheets - Headers don\'t match, but keeping existing data');
      // Don't clear the sheet, just ensure we have the right number of columns
      if (lastCol < headers.length) {
        // Add missing columns
        for (let i = lastCol + 1; i <= headers.length; i++) {
          results.getRange(1, i).setValue(headers[i-1]);
        }
        SpreadsheetApp.flush(); // Force flush to ensure headers are written
      }
    }
  }
  
  // Verify the sheet is properly set up
  const finalLastRow = results.getLastRow();
  const finalLastCol = results.getLastColumn();
  Logger.log('ensureSheets - Final sheet state: %s rows, %s columns', finalLastRow, finalLastCol);
  
  // Verify headers are correct
  if (finalLastRow > 0) {
    const finalHeaders = results.getRange(1,1,1,finalLastCol).getValues()[0];
    Logger.log('ensureSheets - Final headers: %s', finalHeaders.join(' | '));
  }
  
  return results;
}

/* =========================
   ENTRY: startCollection
   - Initializes state and starts trigger-based batched processing.
   ========================= */
function startCollection() {
  try { repairApiKeyIfMissing(); } catch (e) {}
  const existingKey = ensureApiKeyInteractive();
  if (!existingKey) { return; }

  // prepare sheets
  ensureSheets();

  // clear old triggers & progress
  clearTriggers();
  clearProgressOnly();

  // initial state
  const state = {
    phase: 'search',          // 'search' or 'details' or 'done'
    centerIndex: 0,
    keywordIndex: 0,
    next_page_token: null,
    batchProcessedCount: 0,
    lastRunAt: new Date().toISOString()
  };

  setProp('STATE', state);
  // create immediate trigger to start quickly (time-based min 1 minute) OR call continueCollection now for immediate start
  // To avoid hitting timeouts for long immediate run, we'll run continueCollection synchronously once (itself enqueues next trigger if more)
  SpreadsheetApp.getUi().alert('Collection started. The script will run in batches to avoid execution limits. You will be notified of errors via alerts.');
  // Run first batch immediately (synchronous)
  continueCollection();
}

/* =========================
   Continue processing (called by trigger or initial run)
   - Reads state and advances processing
   - Creates a time-based trigger to self if still work to do
   ========================= */
function continueCollection() {
  const ui = SpreadsheetApp.getUi();
  try {
    try { repairApiKeyIfMissing(); } catch (e) {}
    const state = getProp('STATE');
    if (!state) {
      ui.alert('No state found. Start collection first from menu.');
      return;
    }
    const apiKey = ensureApiKeyInteractive();
    if (!apiKey) { ui.alert('API key not found. Please run IndoorGolf > Set API Key to set it.'); return; }
    ensureSheets();
    // route by phase
    if (state.phase === 'search') {
      const finished = processSearchBatch(state, apiKey);
      if (finished) {
        // move to details phase
        state.phase = 'details';
        state.centerIndex = 0;
        state.keywordIndex = 0;
        state.next_page_token = null;
        state.batchProcessedCount = 0;
        setProp('STATE', state);
        // schedule next batch to start details
        createTriggerForContinue();
        ui.alert('Search phase completed. Now starting Details fetch in batches.');
        return;
      } else {
        // still searching: schedule next
        setProp('STATE', state);
        createTriggerForContinue();
        return;
      }
    } else if (state.phase === 'details') {
      const finished = processDetailsBatch(state, apiKey);
      if (finished) {
        state.phase = 'done';
        setProp('STATE', state);
        clearTriggers(); // all done
        ui.alert('All details fetched. Collection finished.');
        return;
      } else {
        setProp('STATE', state);
        createTriggerForContinue();
        return;
      }
    } else if (state.phase === 'done') {
      ui.alert('Collection already finished. If you want to re-run, click Start Collection.');
      return;
    } else {
      ui.alert('Unknown state phase: ' + state.phase);
      return;
    }
  } catch (e) {
    // show error
    try { ui.alert('Error in continueCollection: ' + String(e)); }
    catch (ee) { Logger.log('Error showing alert: ' + ee); }
    // ensure we clear triggers to avoid looping on error unless you want retry
    clearTriggers(); 
    throw e;
  }
}

/* =========================
   processSearchBatch
   - Performs Text Search calls, appends deduped rows to Results sheet
   - Returns true when search complete (no more centers/keywords)
   ========================= */
function processSearchBatch(state, apiKey) {
  const resultsSheet = ensureSheets();
  const headers = resultsSheet.getRange(1,1,1,resultsSheet.getLastColumn()).getValues()[0];
  const existingPlaceIds = getExistingPlaceIds(resultsSheet);

  let processedThisBatch = 0;
  const maxPerBatch = CONFIG.BATCH_SIZE_SEARCH;

  // iterate centers & keywords from saved indices
  for (let ci = state.centerIndex; ci < CONFIG.CENTERS.length; ci++) {
    const center = CONFIG.CENTERS[ci];
    for (let ki = (ci === state.centerIndex ? state.keywordIndex : 0); ki < CONFIG.KEYWORDS.length; ki++) {
      const keyword = CONFIG.KEYWORDS[ki];
      // handle next_page_token from previous run if present
      let nextToken = state.next_page_token || null;
      let page = 0;
      do {
        page++;
        // Build TextSearch request
        const requestConfig = buildTextSearchUrl(apiKey, keyword, center.lat, center.lng, CONFIG.SEARCH_RADIUS_METERS, nextToken);
        const resp = UrlFetchApp.fetch(requestConfig.url, {
          method: requestConfig.method,
          payload: requestConfig.body,
          headers: requestConfig.headers,
          muteHttpExceptions: true
        });
        const js = JSON.parse(resp.getContentText());
        
        // If no results found, try a broader search without type restriction
        if (resp.getResponseCode() === 200 && (!js.places || js.places.length === 0)) {
          Logger.log('No results with type restriction, trying broader search for: %s', keyword);
          const broaderConfig = buildTextSearchUrl(apiKey, keyword, center.lat, center.lng, CONFIG.SEARCH_RADIUS_METERS, nextToken, false);
          const broaderResp = UrlFetchApp.fetch(broaderConfig.url, {
            method: broaderConfig.method,
            payload: broaderConfig.body,
            headers: broaderConfig.headers,
            muteHttpExceptions: true
          });
          const broaderJs = JSON.parse(broaderResp.getContentText());
          
          if (broaderResp.getResponseCode() === 200 && broaderJs.places && broaderJs.places.length > 0) {
            Logger.log('Broader search found %s results for: %s', broaderJs.places.length, keyword);
            // Use the broader search results
            Object.assign(js, broaderJs);
          } else {
            // Try an even simpler search with just the basic term
            Logger.log('Trying simple search for: %s', keyword);
            const simpleConfig = buildSimpleSearchUrl(apiKey, keyword, center.lat, center.lng);
            const simpleResp = UrlFetchApp.fetch(simpleConfig.url, {
              method: simpleConfig.method,
              payload: simpleConfig.body,
              headers: simpleConfig.headers,
              muteHttpExceptions: true
            });
            const simpleJs = JSON.parse(simpleResp.getContentText());
            
            if (simpleResp.getResponseCode() === 200 && simpleJs.places && simpleJs.places.length > 0) {
              Logger.log('Simple search found %s results for: %s', simpleJs.places.length, keyword);
              Object.assign(js, simpleJs);
            }
          }
        }
        
        // Handle new Places API response format
        Logger.log('Processing search for keyword: %s at center: %s, Response code: %s, Places found: %s', 
          keyword, center.name, resp.getResponseCode(), js.places ? js.places.length : 0);
        
        // Debug: Log the full API response to see what we're getting
        Logger.log('Full API Response: %s', JSON.stringify(js, null, 2));
        
        // Additional debugging for search parameters
        Logger.log('Search Parameters - Keyword: %s, Center: %s (%s, %s), Radius: %s', 
          keyword, center.name, center.lat, center.lng, CONFIG.SEARCH_RADIUS_METERS);
        
        if (resp.getResponseCode() === 200 && js.places && js.places.length) {
          Logger.log('Found %s places for keyword %s at center %s', js.places.length, keyword, center.name);
          
          // iterate results
          for (let r of js.places) {
            const pid = r.id; // New API uses 'id' instead of '店舗ID'
            
            // Log each place found
            Logger.log('Processing place: %s (ID: %s)', r.displayName || 'Unknown', pid);
            
            if (existingPlaceIds[pid]) {
              Logger.log('Place %s already exists, skipping', pid);
              continue;
            }
            
            // prepare row aligned with headers
            const rowObj = {
              店舗ID: pid || '',
              店舗名: r.displayName['text'] || '',
              住所: r.formattedAddress || '',
              種別: getPlaceTypeCategory(r.types),
              電話番号: r.nationalPhoneNumber || '',
              WebサイトURL: r.websiteUri || '',
              評価: r.rating || '',
              口コミ数: r.userRatingCount || '',
              // vicinity: r.formattedAddress || '', // Use formattedAddress as vicinity in new API
              // lat: (r.location && r.location.latitude) || '',
              // lng: (r.location && r.location.longitude) || '',
              // opening_hours_weekday_text: '',
              // open_now: (r.currentOpeningHours && typeof r.currentOpeningHours.openNow !== 'undefined') ? String(r.currentOpeningHours.openNow) : '',
              // types: (r.types || []).join(','),
              // url: r.googleMapsUri || '',
              営業ステータス: r.businessStatus || '',
              検索に使った地域名: center.name,
              検索に使ったキーワード: keyword,
              // collected_at: new Date().toISOString(),
              // details_fetched: 'no'
            };
            
            Logger.log('About to add place to sheet: %s (%s) at center %s with keyword %s', rowObj.店舗名, pid, center.name, keyword);
            Logger.log('Row data: %s', JSON.stringify(rowObj, null, 2));
            
            try {
              // Store the data immediately
              appendRowInOrder(resultsSheet, rowObj);
              existingPlaceIds[pid] = true;
              processedThisBatch++;
              
              Logger.log('SUCCESS: Added place to sheet: %s (%s) - Total processed this batch: %s', rowObj.店舗名, pid, processedThisBatch);
            } catch (e) {
              Logger.log('ERROR: Failed to add place %s to sheet: %s', rowObj.店舗名, e.toString());
            }
            if (processedThisBatch >= maxPerBatch) {
              // save current state and return not finished
              state.centerIndex = ci;
              state.keywordIndex = ki;
              state.next_page_token = js.nextPageToken || null;
              state.batchProcessedCount = (state.batchProcessedCount || 0) + processedThisBatch;
              Logger.log('Search batch processed %s rows; pausing to avoid time limits.', processedThisBatch);
              return false;
            }
          } // end for results
        } else if (resp.getResponseCode() === 200 && (!js.places || js.places.length === 0)) {
          // No results found
          Logger.log('No results found for keyword: %s at center: %s', keyword, center.name);
          Logger.log('API Response when no results: %s', JSON.stringify(js, null, 2));
        } else {
          // Handle errors
          Logger.log('API Error - Response Code: %s, Body: %s', resp.getResponseCode(), resp.getContentText());
          if (resp.getResponseCode() === 429 || resp.getResponseCode() === 403) {
            // Rate limit or permission error - wait and retry
            Utilities.sleep(2000);
            continue;
          } else {
            // show error and stop
            SpreadsheetApp.getUi().alert('TextSearch API error - Code: ' + resp.getResponseCode() + '\nMessage body: ' + resp.getContentText());
            throw new Error('TextSearch API error: ' + resp.getResponseCode());
          }
        }

        // handle next_page_token (new API uses nextPageToken)
        if (js.nextPageToken) {
          // set for next iteration and wait before calling again
          nextToken = js.nextPageToken;
          state.next_page_token = nextToken;
          Utilities.sleep(CONFIG.NEXT_PAGE_WAIT_MS);
        } else {
          nextToken = null;
          state.next_page_token = null;
        }
      } while (nextToken);

      // finished the keyword for this center; reset next_page_token
      state.next_page_token = null;
      state.keywordIndex = ki + 1;
    } // end for keywords
    // finished keywords for this center
    state.keywordIndex = 0;
    state.centerIndex = ci + 1;
  } // end for centers

  // finished all centers & keywords
  state.batchProcessedCount = (state.batchProcessedCount || 0) + processedThisBatch;
  Logger.log('Search phase completed. Total rows added this run: %s', processedThisBatch);
  
  // Final verification - check if any data was actually written to the sheet
  const finalSheet = ensureSheets();
  const finalRowCount = finalSheet.getLastRow();
  Logger.log('Final sheet status: %s rows total', finalRowCount);
  
  if (finalRowCount <= 1) {
    Logger.log('WARNING: No data rows found in sheet after search completion!');
    
    // Try to add some test data to verify the sheet writing works
    Logger.log('Adding test data to verify sheet writing works...');
    const testData = {
      店舗ID: 'TEST_NO_RESULTS',
      店舗名: 'Test - No Golf Courses Found',
      住所: 'Test Address - No Results Found',
      種別: 'アウトドア',
      電話番号: '',
      WebサイトURL: '',
      評価: '0',
      口コミ数: '0',
      // vicinity: 'Test Address - No Results Found',
      // lat: '35.6762',
      // lng: '139.6503',
      // opening_hours_weekday_text: '',
      // open_now: '',
      // types: 'test',
      // url: '',
      営業ステータス: 'TEST',
      検索に使った地域名: 'Test',
      検索に使ったキーワード: 'test',
      // collected_at: new Date().toISOString(),
      // details_fetched: 'no'
    };
    
    try {
      appendRowInOrder(finalSheet, testData);
      Logger.log('Added test row to verify sheet writing works');
    } catch (e) {
      Logger.log('Failed to add test row: %s', e.toString());
    }
    
    SpreadsheetApp.getUi().alert('WARNING: Search completed but no golf course data was found. Added test row to verify sheet writing. Check the logs for API response details.');
  } else {
    Logger.log('SUCCESS: Found %s data rows in sheet', finalRowCount - 1);
  }
  
  return true;
}

/* =========================
   processDetailsBatch
   - Fetches Place Details for rows with details_fetched === 'no'
   - Updates rows with phone, website, opening hours, open_now, url, 営業ステータス
   - Returns true when all rows updated
   ========================= */
function processDetailsBatch(state, apiKey) {
  const sheet = ensureSheets();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  Logger.log('Details batch - Sheet has %s rows total', values.length);
  
  if (values.length <= 1) {
    Logger.log('Details batch - No data rows found, nothing to do');
    return true;
  }
  
  const header = values[0];
  const rows = values.slice(1); // omit header
  
  Logger.log('Details batch - Headers: %s', header.join(', '));
  
  // Build column map
  const colIndex = {};
  header.forEach((h, i) => colIndex[h] = i);
  
  Logger.log('Details batch - Column index: %s', JSON.stringify(colIndex));

  // Find rows that need details (details_fetched === 'no')
  const rowsNeeding = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const detailsFetched = row[colIndex['details_fetched']] || '';
    const pid = row[colIndex['店舗ID']];
    if ((detailsFetched === 'no' || detailsFetched === '') && pid) {
      rowsNeeding.push({ rowIndex: i + 2, 店舗ID: pid }); // +2 because header and 1-index
    }
  }

  Logger.log('Details batch - Found %s rows needing details', rowsNeeding.length);

  if (rowsNeeding.length === 0) {
    Logger.log('No rows need details. Details phase complete.');
    return true;
  }

  // process up to BATCH_SIZE_DETAILS
  const toProcess = rowsNeeding.slice(0, CONFIG.BATCH_SIZE_DETAILS);
  Logger.log('Details batch - Processing %s rows', toProcess.length);
  
  for (let item of toProcess) {
    try {
      Logger.log('Details batch - Fetching details for 店舗ID: %s (row %s)', item.店舗ID, item.rowIndex);
      const details = fetchPlaceDetails(apiKey, item.店舗ID);
      Logger.log('Details batch - Got details for %s: %s', item.店舗ID, JSON.stringify(details));
      
      // write back into sheet at item.rowIndex
      const writeRange = sheet.getRange(item.rowIndex, 1, 1, header.length);
      // read current row to preserve existing search data
      const currentRow = sheet.getRange(item.rowIndex, 1, 1, header.length).getValues()[0];

      // set fields by header name if present
      if (colIndex['電話番号'] !== undefined) currentRow[colIndex['電話番号']] = details.電話番号 || currentRow[colIndex['電話番号']] || '';
      if (colIndex['WebサイトURL'] !== undefined) currentRow[colIndex['WebサイトURL']] = details.WebサイトURL || currentRow[colIndex['WebサイトURL']] || '';
      // if (colIndex['opening_hours_weekday_text'] !== undefined) currentRow[colIndex['opening_hours_weekday_text']] = details.opening_hours_weekday_text || currentRow[colIndex['opening_hours_weekday_text']] || '';
      // if (colIndex['open_now'] !== undefined) currentRow[colIndex['open_now']] = (typeof details.open_now !== 'undefined') ? String(details.open_now) : currentRow[colIndex['open_now']] || '';
      if (colIndex['住所'] !== undefined) currentRow[colIndex['住所']] = details.住所 || currentRow[colIndex['住所']] || '';
      if (colIndex['評価'] !== undefined) currentRow[colIndex['評価']] = details.評価 || currentRow[colIndex['評価']] || '';
      if (colIndex['口コミ数'] !== undefined) currentRow[colIndex['口コミ数']] = details.口コミ数 || currentRow[colIndex['口コミ数']] || '';
      // if (colIndex['url'] !== undefined) currentRow[colIndex['url']] = details.url || currentRow[colIndex['url']] || '';
      // if (colIndex['types'] !== undefined) currentRow[colIndex['types']] = details.types || currentRow[colIndex['types']] || '';
      if (colIndex['種別'] !== undefined) currentRow[colIndex['種別']] = details.種別 || currentRow[colIndex['種別']] || '';
      if (colIndex['営業ステータス'] !== undefined) currentRow[colIndex['営業ステータス']] = details.営業ステータス || currentRow[colIndex['営業ステータス']] || '';
      // if (colIndex['collected_at'] !== undefined && !currentRow[colIndex['collected_at']]) currentRow[colIndex['collected_at']] = new Date().toISOString();
      // if (colIndex['details_fetched'] !== undefined) currentRow[colIndex['details_fetched']] = 'yes';
      
      // write back
      writeRange.setValues([currentRow]);
      Logger.log('Details batch - Successfully updated row %s with details for %s', item.rowIndex, item.店舗ID);
    } catch (e) {
      Logger.log('Error fetching details for %s: %s', item.店舗ID, e);
      // mark row with error in details_fetched column if possible
      try {
        const idx = colIndex['details_fetched'];
        if (idx !== undefined) {
          sheet.getRange(item.rowIndex, idx + 1).setValue('error');
          Logger.log('Marked row %s as error for 店舗ID %s', item.rowIndex, item.店舗ID);
        }
      } catch (ee) { Logger.log('Also failed to mark error: ' + ee); }
    }
    Utilities.sleep(CONFIG.DETAILS_SLEEP_MS);
  }

  // if there are more rows needing details, return false to schedule next trigger
  if (rowsNeeding.length > toProcess.length) {
    Logger.log('Processed %s details; %s remaining. Scheduling next batch.', toProcess.length, rowsNeeding.length - toProcess.length);
    return false;
  } else {
    Logger.log('All details fetched for all rows.');
    return true;
  }
}

/* =========================
   Helper: build Simple Search URL (fallback method)
   ========================= */
function buildSimpleSearchUrl(apiKey, query, lat, lng) {
  const base = 'https://places.googleapis.com/v1/places:searchText';
  const params = [];
  params.push('key=' + encodeURIComponent(apiKey));
  
  // Corrected request body format for Google Places API (New)
  const requestBody = {
    textQuery: query,
    locationBias: {
      circle: {
        center: {
          latitude: parseFloat(lat),
          longitude: parseFloat(lng)
        },
        radius: 50000 // 50km radius for broader search
      }
    },
    languageCode: 'ja',
    maxResultCount: 20,
    regionCode: 'JP'
  };
  
  Logger.log('Simple Search Request Body: %s', JSON.stringify(requestBody, null, 2));
  
  return {
    url: base + '?' + params.join('&'),
    method: 'POST',
    body: JSON.stringify(requestBody),
    headers: {
      'Content-Type': 'application/json',
      'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.location,places.rating,places.userRatingCount,places.types,places.businessStatus,places.currentOpeningHours,places.nationalPhoneNumber,places.websiteUri,places.googleMapsUri'
    }
  };
}

/* =========================
   Helper: build Text Search URL (New Places API)
   ========================= */
function buildTextSearchUrl(apiKey, query, lat, lng, radius, nextPageToken, includeTypeRestriction = true) {
  const base = 'https://places.googleapis.com/v1/places:searchText';
  const params = [];
  params.push('key=' + encodeURIComponent(apiKey));
  
  // Build request body for POST request - corrected format
  const requestBody = {
    textQuery: query,
    locationBias: {
      circle: {
        center: {
          latitude: parseFloat(lat),
          longitude: parseFloat(lng)
        },
        radius: parseInt(radius)
      }
    },
    languageCode: 'ja',
    maxResultCount: 20,
    regionCode: 'JP'
  };
  
  // Only add type restriction if requested (for broader search)
  if (includeTypeRestriction) {
    requestBody.includedType = 'golf_course';
  }
  
  if (nextPageToken) {
    requestBody.pageToken = nextPageToken;
  }
  
  // Log the request for debugging
  Logger.log('API Request Body: %s', JSON.stringify(requestBody, null, 2));
  
  return {
    url: base + '?' + params.join('&'),
    method: 'POST',
    body: JSON.stringify(requestBody),
    headers: {
      'Content-Type': 'application/json',
      'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.location,places.rating,places.userRatingCount,places.types,places.businessStatus,places.currentOpeningHours,places.nationalPhoneNumber,places.websiteUri,places.googleMapsUri'
    }
  };
}

/* =========================
   Helper: fetchPlaceDetails (New Places API)
   - Requests useful fields
   ========================= */
function fetchPlaceDetails(apiKey, placeId) {
  const base = 'https://places.googleapis.com/v1/places/' + encodeURIComponent(placeId);
  const params = [];
  params.push('key=' + encodeURIComponent(apiKey));
  params.push('languageCode=ja');
  
  const url = base + '?' + params.join('&');
  const options = {
    method: 'GET',
    headers: {
      'X-Goog-FieldMask': 'id,displayName,formattedAddress,location,rating,userRatingCount,types,businessStatus,currentOpeningHours,regularOpeningHours,nationalPhoneNumber,internationalPhoneNumber,websiteUri,googleMapsUri'
    },
    muteHttpExceptions: true
  };
  
  const resp = UrlFetchApp.fetch(url, options);
  const js = JSON.parse(resp.getContentText());
  
  if (resp.getResponseCode() !== 200) {
    Logger.log('Place Details error for %s: %s', placeId, resp.getResponseCode());
    throw new Error('Place Details error: ' + resp.getResponseCode());
  }
  
  return {
    住所: js.formattedAddress || '',
    電話番号: js.nationalPhoneNumber || js.internationalPhoneNumber || '',
    WebサイトURL: js.websiteUri || '',
    // opening_hours_weekday_text: (js.regularOpeningHours && js.regularOpeningHours.weekdayDescriptions) ? js.regularOpeningHours.weekdayDescriptions.join(' | ') : '',
    // open_now: (js.currentOpeningHours && typeof js.currentOpeningHours.openNow !== 'undefined') ? js.currentOpeningHours.openNow : '',
    評価: js.rating || '',
    口コミ数: js.userRatingCount || '',
    // url: js.googleMapsUri || '',
    types: (js.types || []).join(','),
    営業ステータス: js.businessStatus || ''
  };
}

/* =========================
   Helpers: read/write sheet rows
   ========================= */
function getExistingPlaceIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; // only header
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const map = {};
  for (let r of data) {
    if (r[0]) map[r[0]] = true;
  }
  return map;
}
function appendRowInOrder(sheet, objRow) {
  try {
    
    if (objRow['検索に使った地域名'] !== "東京" && objRow['検索に使った地域名'] !== "神奈川") {
      return;
    }
    Logger.log('appendRowInOrder - Starting with object: %s', JSON.stringify(objRow, null, 2));
    
    // ensure columns order
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    Logger.log('appendRowInOrder - Headers: %s', headers.join(', '));
    Logger.log('appendRowInOrder - Object keys: %s', Object.keys(objRow).join(', '));
    
    const out = headers.map(h => {
      const value = objRow[h] || '';
      Logger.log('appendRowInOrder - Mapping %s: %s', h, value);
      return value;
    });
    
    Logger.log('appendRowInOrder - Final row data: %s', out.join(' | '));
    
    // Check if sheet is valid
    if (!sheet || !sheet.getRange) {
      throw new Error('Invalid sheet object');
    }
    
    // Force a flush to ensure data is written immediately
    Logger.log('appendRowInOrder - About to call sheet.appendRow');
    sheet.appendRow(out);
    Logger.log('appendRowInOrder - Called sheet.appendRow, now flushing');
    SpreadsheetApp.flush();
    Logger.log('appendRowInOrder - Flush completed');
    
    // Verify the row was added
    const lastRow = sheet.getLastRow();
    Logger.log('appendRowInOrder - Sheet now has %s rows', lastRow);
    
    if (lastRow > 0) {
      const addedRow = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
      Logger.log('appendRowInOrder - Verified row %s: %s', lastRow, addedRow.join(' | '));
      
      // Additional verification - check if the data actually exists
      const actualData = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
      const hasData = actualData.some(cell => cell && cell.toString().trim() !== '');
      if (!hasData) {
        Logger.log('WARNING: Row %s appears to be empty after append', lastRow);
        throw new Error('Row appears empty after append');
      } else {
        Logger.log('SUCCESS: Row %s contains data: %s', lastRow, actualData.join(' | '));
      }
    } else {
      Logger.log('ERROR: No rows in sheet after append');
      throw new Error('No rows in sheet after append');
    }
    
  } catch (e) {
    Logger.log('appendRowInOrder ERROR: %s', e.toString());
    Logger.log('appendRowInOrder ERROR stack: %s', e.stack);
    throw e;
  }
}

// function appendRowInOrder(sheet, objRow) {
//   try {
//     Logger.log('appendRowInOrder - Starting with object: %s', JSON.stringify(objRow, null, 2));
    
//     const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
//     Logger.log('appendRowInOrder - Headers: %s', headers.join(', '));

//     const out = headers.map(h => objRow[h] || '');

//     // ▼▼ ここから追加部分 ▼▼
//     const region = objRow["検索に使用した地域名"]; // オブジェクトの中から地域名を取得
//     if (region !== "東京" && region !== "神奈川") {
//       Logger.log(`地域名「${region}」は対象外のためスキップしました。`);
//       return; // 処理を終了（シートに書き込まない）
//     }
//     // ▲▲ 追加部分ここまで ▲▲

//     Logger.log('appendRowInOrder - Final row data: %s', out.join(' | '));

//     if (!sheet || !sheet.getRange) throw new Error('Invalid sheet object');
    
//     sheet.appendRow(out);
//     SpreadsheetApp.flush();

//     const lastRow = sheet.getLastRow();
//     const addedRow = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
//     Logger.log('SUCCESS: Row %s contains data: %s', lastRow, addedRow.join(' | '));

//   } catch (e) {
//     Logger.log('appendRowInOrder ERROR: %s', e.toString());
//     throw e;
//   }
// }

/* =========================
   Trigger helpers
   ========================= */
function createTriggerForContinue() {
  // remove existing triggers for continueCollection to avoid duplicates
  clearTriggersForFunction('continueCollection');
  ScriptApp.newTrigger('continueCollection').timeBased().after(CONFIG.TRIGGER_RETRY_MINUTES * 60 * 1000).create();
}
function clearTriggersForFunction(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });
}
function clearTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
}
function clearTriggersAndProgress() {
  clearTriggers();
  clearProgressOnly();
  try { repairApiKeyIfMissing(); } catch (e) {}
  SpreadsheetApp.getUi().alert('Cleared triggers and saved progress.');
}

/* =========================
   CSV Save: export Results sheet to Drive as CSV
   ========================= */
function saveAsCSV() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.RESULTS_SHEET);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Results sheet not found.');
      return;
    }
    const csv = convertRangeToCsvFile_(sheet);
    const fileName = 'IndoorGolf_Results_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
    const file = DriveApp.createFile(fileName, csv, MimeType.CSV);
    SpreadsheetApp.getUi().alert('CSV saved to Drive as: ' + file.getName());
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error saving CSV: ' + e);
    throw e;
  }
}
function convertRangeToCsvFile_(sheet) {
  const data = sheet.getDataRange().getValues();
  let csv = '';
  data.forEach(function(rowArray) {
    const row = rowArray.map(cell => {
      if (cell === null || typeof cell === 'undefined') return '';
      // escape double quotes
      const v = String(cell).replace(/"/g, '""');
      // wrap if contains comma/newline/quote
      if (v.indexOf(',') >= 0 || v.indexOf('\n') >= 0 || v.indexOf('"') >= 0) return '"' + v + '"';
      return v;
    }).join(',');
    csv += row + '\r\n';
  });
  return csv;
}

/* =========================
   Utility: remove all internal props (for debugging)
   ========================= */
function debug_clearProps() {
  clearAllIndoorGolfProps();
  Logger.log('Cleared IndoorGolf properties.');
}

/* =========================
   End of script
   ========================= */
