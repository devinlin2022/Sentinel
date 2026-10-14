/**
 * Serves the primary HTML file for the web app.
 * @param {Object} e The event parameter from the web request.
 * @return {HtmlOutput} The HTML page to be rendered for the user.
 */
function doGet(e) {
  // The name 'index' here MUST EXACTLY match the name of your HTML file (index.html).
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Sentinel Dashboard'); 
}

function getCombinedData() {
  const SHEET_ID = '1enQAz_hyKFOFndS8cX1veNt4M8svsmoGzpc_Eiagu5Q';

  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    // 1. Process data from the 'Data' sheet
    const shipmentPayload = processShipmentData(spreadsheet);
    if (shipmentPayload.error) throw new Error(shipmentPayload.error);
    
    const predictivePayload = processPredictiveData(spreadsheet, shipmentPayload.laneIdToRouteMap); 

    // Check for error in the second sheet
    if (predictivePayload.error) throw new Error(predictivePayload.error);

    // --- Filter Sorting (Done separately for D1/2 and D3) ---
    // Sort shipment filter options
    for (const key in shipmentPayload.filters) {
        shipmentPayload.filters[key].sort((a,b) => {
            if (key === 'atdMonth' || key === 'rgMonth') return String(a).localeCompare(String(b));
            const isNumericA = !isNaN(parseFloat(a)) && isFinite(a);
            const isNumericB = !isNaN(parseFloat(b)) && isFinite(b);
            if (isNumericA && isNumericB) return a - b;
            return String(a).localeCompare(String(b), undefined, {numeric: true});
        });
    }
    // Sort predictive filter options
    for (const key in predictivePayload.filters) {
        predictivePayload.filters[key].sort((a,b) => {
            if (key === 'atdMonth' || key === 'rgMonth') return String(a).localeCompare(String(b));
            const isNumericA = !isNaN(parseFloat(a)) && isFinite(a);
            const isNumericB = !isNaN(parseFloat(b)) && isFinite(b);
            if (isNumericA && isNumericB) return a - b;
            return String(a).localeCompare(String(b), undefined, {numeric: true});
        });
    }

    // Assemble the final payload for the frontend.
    const finalPayload = {
      shipments: shipmentPayload.shipments,
      predictive: predictivePayload, 
      filters: shipmentPayload.filters, 
      dateInfo: getCurrentDateInfo()
    };
    
    return JSON.stringify(finalPayload);

  } catch (e) {
    Logger.log(`CRITICAL ERROR in getCombinedData: ${e.stack}`);
    return JSON.stringify({ error: `An error occurred on the server: ${e.message}` });
  }
}

/**
 * Processes data from the 'Data' sheet for Dashboards 1, 2.
 * @param {Spreadsheet} spreadsheet The active spreadsheet object.
 * @return {Object} An object containing shipments data, filter options, and laneIdToRouteMap.
 */
function processShipmentData(spreadsheet) {
    const SHEET_NAME = 'Data';
    // Headers based on the 'Data.csv' file structure - VERIFIED
    const HEADERS = {
      container: 'Container',          // Col A
      polCity: 'POL City',             // Col K
      podCity: 'POD City',             // Col M
      vesselEtd: 'Vessel ETD',         // Col W
      vesselAtd: 'Vessel ATD',         // Col Y
      vesselEta: 'Vessel ETA',         // Col AA
      vesselAta: 'Vessel ATA',         // Col AF
      atdMonth: 'ATD Month',           // Col AT
      regionHub: 'Region/Hub',         // Col AU
      finalDestCountry: 'Final Dest. Country', // Col AW
      laneId: 'Lane ID',               // Col BN
      rgRate: 'RG Rate',               // Col BX
      budRate: 'BUD Rate',             // Col BZ
      xenetaRate: 'Xeneta Rate',       // Col CA
      actualSaving: 'Actual Saving(ATD month RG- BUD Rate)', // Col CB
      status: 'Gnosis Milestone',      // Col CC
      rgTransitTime: 'RG Transit Time',// Col CD
      etdWeek: 'Vessel ETD Week',      // Col CI
      etaWeek: 'Vessel ETA Week',      // Col CJ
      departureDelay: 'ATD Vs ETD Delay', // Col CK
      arrivalDelay: 'ATA Vs ETA Delay',   // Col CL
      potentialSaving: 'Potential Saving(ATD month RG- Xeneta)' // Col CM
    };
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: `Sheet named "${SHEET_NAME}" was not found.` };
   
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { shipments: [], filters: {}, laneIdToRouteMap: {} };
    
    const headerMap = values[0].reduce((acc, header, index) => { acc[header.trim()] = index; return acc; }, {});
    
    // Validate all required headers exist
    for (const key in HEADERS) {
        if (headerMap[HEADERS[key]] === undefined) {
            return { error: `Required column header not found in '${SHEET_NAME}': "${HEADERS[key]}". Please check spreadsheet.`};
        }
    }
   
    // Filter options for Dashboard 1 and 2
    const filterOptions = {
        regionHub: new Set(), finalDestCountry: new Set(), polCity: new Set(), podCity: new Set(),
        status: new Set(), etdWeek: new Set(), etaWeek: new Set(), departureDelay: new Set(),
        arrivalDelay: new Set(), atdMonth: new Set(), laneId: new Set()
    };
    
    const laneIdToRouteMap = {};
   
    // NEW: Define the date filter: 08/01/2025
    const filterDate = new Date('2025-10-01T00:00:00Z');
    const filterTimestamp = filterDate.getTime();
    
    const processedData = values.slice(1).map(row => {
        const get = (headerName) => row[headerMap[headerName]];
       
        // --- Date Filtering ---
        const vesselEtdValue = get(HEADERS.vesselEtd);
        const vesselEtd = parseDate(vesselEtdValue);

        // Skip row if ETD is invalid or before the filter date
        if (!vesselEtd || isNaN(vesselEtd.getTime()) || vesselEtd.getTime() < filterTimestamp) {
            return null; // Skip this row
        }

        const vesselAtd = parseDate(get(HEADERS.vesselAtd));
        const vesselEta = parseDate(get(HEADERS.vesselEta));
        const vesselAta = parseDate(get(HEADERS.vesselAta));
        
        let actualTransitTime = null;
        if (vesselAta && !isNaN(vesselAta.getTime()) && vesselAtd && !isNaN(vesselAtd.getTime())) {
            const timeDiff = vesselAta.getTime() - vesselAtd.getTime();
            if (timeDiff >= 0) actualTransitTime = Math.floor(timeDiff / (1000 * 3600 * 24)); // In days
        }

        const laneIdValue = get(HEADERS.laneId);
        const polCityValue = get(HEADERS.polCity);
        const podCityValue = get(HEADERS.podCity);

        const shipment = {
            container: get(HEADERS.container),
            vesselEtd: formatDateForTable(vesselEtd), 
            vesselAtd: formatDateForTable(vesselAtd), 
            vesselEta: formatDateForTable(vesselEta),
            vesselAta: formatDateForTable(vesselAta), 
            actualOceanTransitTime: actualTransitTime,
            status: get(HEADERS.status),
            // NEW: Use the formatISOYearWeek helper to construct "YYYYWww" for charts
            etdWeek: formatISOYearWeek(vesselEtd) || get(HEADERS.etdWeek),
            etaWeek: formatISOYearWeek(vesselEta) || get(HEADERS.etaWeek),
            departureDelayCategory: get(HEADERS.departureDelay),
            arrivalDelayCategory: get(HEADERS.arrivalDelay),
            laneId: laneIdValue,
            polCity: polCityValue,
            podCity: podCityValue,
            rgTransitTime: parseFloat(get(HEADERS.rgTransitTime)) || null,
            regionHub: get(HEADERS.regionHub),
            finalDestCountry: get(HEADERS.finalDestCountry),
            atdMonth: parseInt(get(HEADERS.atdMonth), 10) || null,
            rgRate: parseFloat(get(HEADERS.rgRate)) || 0,
            budRate: parseFloat(get(HEADERS.budRate)) || 0,
            actualSaving: parseFloat(get(HEADERS.actualSaving)) || 0,
            xenetaRate: parseFloat(get(HEADERS.xenetaRate)) || null,
            potentialSaving: parseFloat(get(HEADERS.potentialSaving)) || 0,
        };

        // Populate route map
        if (laneIdValue && !laneIdToRouteMap[laneIdValue] && polCityValue && podCityValue) {
            laneIdToRouteMap[laneIdValue] = `${polCityValue} -> ${podCityValue}`;
        }

        // Populate filter options *after* date filter has passed
        Object.keys(filterOptions).forEach(key => {
            const dataKey = key === 'departureDelay' ? 'departureDelayCategory' : key === 'arrivalDelay' ? 'arrivalDelayCategory' : key;
            const value = shipment[dataKey];
            if (value !== null && value !== undefined && value !== "") {
                filterOptions[key].add(value);
            }
        });
        
        return shipment;
    }).filter(Boolean); // Filter out the null rows skipped by the date filter

    const sortedFilters = {};
    for (const key in filterOptions) {
        sortedFilters[key] = Array.from(filterOptions[key]);
    }
   
    return { shipments: processedData, filters: sortedFilters, laneIdToRouteMap: laneIdToRouteMap };
}

/**
 * Processes data from 'Accumulated RG' and generates future projections for Dashboard 3.
 * @param {Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {Object} laneIdToRouteMap A map from 'Lane ID' (from Data sheet) to 'POL -> POD' string.
 * @return {Object} An object containing { predictiveData, filters, sortedLaneIds }.
 */
function processPredictiveData(spreadsheet, laneIdToRouteMap = {}) {
    const SHEET_NAME = 'Accumulated RG';
    const HEADERS = {
        lane: 'Lane',
        start: 'Start',
        ff: 'FF', 
        originCountry: 'Origin Country', 
        pol: 'Port of Load',
        destinationSubRegion: 'Destination Sub Region/Hub', 
        pod: 'Port of Discharge',
        oceanFreight: 'Ocean Freight Rate (USD)',
        baf: 'BAF (USD)',
        xenetaRate: 'Xeneta Rate',
        annualVolume: '_2023-2024 Volume (container) Annual'
    };
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: `Sheet named "${SHEET_NAME}" was not found.` };
   
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { predictiveData: [], filters: {}, sortedLaneIds: [] };
    
    const headerMap = values[0].reduce((acc, header, index) => { acc[header.trim()] = index; return acc; }, {});
    
    // Validate required headers
    for (const key in HEADERS) {
        if (headerMap[HEADERS[key]] === undefined) {
            Logger.log(`Missing header in '${SHEET_NAME}': "${HEADERS[key]}"`);
            throw new Error(`Required column header not found in '${SHEET_NAME}': "${HEADERS[key]}".`);
        }
    }

    const monthlyDataTemp = {};
    const allMonths = new Set();
    const allLanes = new Set();
    const allFF = new Set();
    const allOriginCountries = new Set();
    const allDestSubRegions = new Set();
    const allPOL = new Set();
    const allPOD = new Set();
    const allGnosisRoutes = new Set();
    
    values.slice(1).forEach(row => {
        const get = (headerName) => row[headerMap[headerName]];
        const startDate = parseDate(get(HEADERS.start));
        const laneId = get(HEADERS.lane); 

        if (!startDate || !laneId) return; 

        const rgMonth = startDate.getFullYear() * 100 + (startDate.getMonth() + 1); 
        const key = `${laneId}_${rgMonth}`;
       
        const ofRate = parseFloat(get(HEADERS.oceanFreight)) || 0;
        const bafRate = parseFloat(get(HEADERS.baf)) || 0;
        const xeneta = parseFloat(get(HEADERS.xenetaRate)) || null; 
        const pol = get(HEADERS.pol);
        const pod = get(HEADERS.pod);
        const annualVolume = parseFloat(String(get(HEADERS.annualVolume)).replace(/,/g, '')) || 0;
        const ff = get(HEADERS.ff);
        const originCountry = get(HEADERS.originCountry);
        const destinationSubRegion = get(HEADERS.destinationSubRegion);
        const gnosisRoute = laneIdToRouteMap[laneId] || `${pol} -> ${pod}` || laneId;

        if (!monthlyDataTemp[key]) {
            monthlyDataTemp[key] = {
                laneId: laneId, rgMonth: rgMonth, pol: pol, pod: pod,
                ff: ff, originCountry: originCountry, destinationSubRegion: destinationSubRegion,
                gnosisRoute: gnosisRoute, annualVolume: annualVolume, 
                ofBafRates: [], xenetaRates: [], count: 0
            };
        }
        monthlyDataTemp[key].ofBafRates.push(ofRate + bafRate);
        if (xeneta !== null) { 
          monthlyDataTemp[key].xenetaRates.push(xeneta);
        }
        monthlyDataTemp[key].count++;

        allMonths.add(rgMonth);
        allLanes.add(laneId);
        if (ff) allFF.add(ff);
        if (originCountry) allOriginCountries.add(originCountry);
        if (destinationSubRegion) allDestSubRegions.add(destinationSubRegion);
        if (pol) allPOL.add(pol);
        if (pod) allPOD.add(pod);
        if (gnosisRoute) allGnosisRoutes.add(gnosisRoute);
    });
    
    const historicalDataByLane = {};
    for (const key in monthlyDataTemp) {
        const entry = monthlyDataTemp[key];
        const avgOfBaf = entry.ofBafRates.reduce((a, b) => a + b, 0) / entry.count;
        const avgXeneta = entry.xenetaRates.length > 0 ? entry.xenetaRates.reduce((a, b) => a + b, 0) / entry.xenetaRates.length : null;
        
        const dataRow = {
            laneId: entry.laneId, rgMonth: entry.rgMonth, pol: entry.pol, pod: entry.pod,
            ff: entry.ff, originCountry: entry.originCountry, destinationSubRegion: entry.destinationSubRegion,
            gnosisRoute: entry.gnosisRoute, annualVolume: entry.annualVolume,
            ofBaf: avgOfBaf,
            xenetaRate: avgXeneta,
            projected: false
        };

        if (!historicalDataByLane[entry.laneId]) historicalDataByLane[entry.laneId] = [];
        historicalDataByLane[entry.laneId].push(dataRow);
    }

    const today = new Date();
    const currentMonth = today.getMonth(); 
    const currentYear = today.getFullYear();
    const finalProjectedData = [];
    const movingAvgWindow = 3; 
    const laneSortData = [];
    
    for (const laneId in historicalDataByLane) {
        const laneData = historicalDataByLane[laneId].sort((a, b) => a.rgMonth - b.rgMonth);
        if (laneData.length === 0) continue;
        
        for(let i=0; i < laneData.length; i++){
            if(i >= movingAvgWindow - 1 &&
               laneData[i-2].xenetaRate !== null &&
               laneData[i-1].xenetaRate !== null &&
               laneData[i].xenetaRate !== null) {
                let sum = 0;
                let count = 0; 
                for(let j=0; j < movingAvgWindow; j++){
                        sum += laneData[i-j].xenetaRate;
                        count++;
                }
                laneData[i].xenetaMovingAvg = sum / count;
            } else {
                laneData[i].xenetaMovingAvg = null;
            }
        }
       
        const lastDataPoint = laneData[laneData.length - 1];
        laneSortData.push({ laneId: laneId, volume: lastDataPoint.annualVolume });
        finalProjectedData.push(...laneData);
        
        const lastFewXeneta = laneData.slice(-movingAvgWindow).map(d => d.xenetaRate);
        while(lastFewXeneta.length < movingAvgWindow) {
            lastFewXeneta.unshift(null);
        }

        for (let i = 1; i <= 6; i++) {
            const projectionDate = new Date(currentYear, currentMonth + i, 1);
            const projectedMonth = projectionDate.getFullYear() * 100 + (projectionDate.getMonth() + 1);
            
            const predictedXeneta = lastDataPoint.xenetaRate;
            lastFewXeneta.shift();
            lastFewXeneta.push(predictedXeneta);
           
            let projectedMovingAvg = null;
            const validPoints = lastFewXeneta.filter(v => v !== null && !isNaN(v));
            if (validPoints.length === movingAvgWindow) {
                projectedMovingAvg = validPoints.reduce((a, b) => a + b, 0) / movingAvgWindow;
            }
           
            finalProjectedData.push({
                laneId: laneId,
                rgMonth: projectedMonth,
                ofBaf: lastDataPoint.ofBaf, 
                xenetaRate: predictedXeneta, 
                annualVolume: lastDataPoint.annualVolume, 
                pol: lastDataPoint.pol, 
                pod: lastDataPoint.pod,
                ff: lastDataPoint.ff,
                originCountry: lastDataPoint.originCountry,
                destinationSubRegion: lastDataPoint.destinationSubRegion,
                gnosisRoute: lastDataPoint.gnosisRoute,
                projected: true, 
                xenetaMovingAvg: projectedMovingAvg 
            });
            allMonths.add(projectedMonth);
        }
    }
   
    const sortedLaneIds = laneSortData.sort((a, b) => b.volume - a.volume).map(l => l.laneId);
    
    const filters = {
        laneId: Array.from(allLanes),
        atdMonth: Array.from(allMonths),
        ff: Array.from(allFF),
        originCountry: Array.from(allOriginCountries),
        destinationSubRegion: Array.from(allDestSubRegions),
        pol: Array.from(allPOL), 
        pod: Array.from(allPOD), 
        gnosisRoute: Array.from(allGnosisRoutes)
    };
    
    return {
      predictiveData: finalProjectedData,
      filters: filters,
      sortedLaneIds: sortedLaneIds 
    };
}

// --- HELPER FUNCTIONS ---

/**
 * Formats a date object into ISO Year-Week string (e.g. 2026W01).
 * @param {Date|null} date The Date object (assumed to be UTC).
 * @return {string|null} The formatted year-week string or null.
 */
function formatISOYearWeek(date) {
    if (!date || isNaN(date.getTime())) return null;
    const d = new Date(date.getTime());
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const year = d.getUTCFullYear();
    const yearStart = new Date(Date.UTC(year, 0, 1));
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${year}W${String(weekNo).padStart(2, '0')}`;
}

function parseDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date) {
      return new Date(Date.UTC(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()));
  }

  if (typeof dateValue === 'number' && dateValue > 0) { 
      if (dateValue > 1 && dateValue < 100000) {
        try {
          const epoch = Date.UTC(1899, 11, 30);
          const millisPerDay = 86400 * 1000;
          const dateMillis = epoch + (dateValue * millisPerDay);
          const resultDate = new Date(dateMillis);
           if (!isNaN(resultDate.getTime())) {
                return new Date(Date.UTC(resultDate.getUTCFullYear(), resultDate.getUTCMonth(), resultDate.getUTCDate()));
           }
        } catch (e) {
           Logger.log(`Could not parse serial number ${dateValue} as date.`);
        }
      }
      return null;
  }

  if (typeof dateValue === 'string') {
    let dateStr = dateValue.trim();
    let parts;

    if (/^\d{4}-\d{1,2}-\d{1,2}/.test(dateStr)) {
        parts = dateStr.split(' ')[0].split('-');
        const date = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
        if (!isNaN(date.getTime())) return date;
    }

    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(dateStr)) {
        parts = dateStr.split(' ')[0].split('/');
        const date = new Date(Date.UTC(parseInt(parts[2], 10), parseInt(parts[0], 10) - 1, parseInt(parts[1], 10)));
        if (!isNaN(date.getTime())) return date;
    }
     
     if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(dateStr)) { 
        parts = dateStr.split(' ')[0].split('/');
        if (parseInt(parts[0], 10) > 12 && parseInt(parts[1], 10) <= 12) {
             const date = new Date(Date.UTC(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10)));
             if (!isNaN(date.getTime())) return date;
        }
    }

    try {
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
            return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        }
    } catch(e) {
         Logger.log(`Native date parse failed for string: ${dateStr}`);
    }
  }

  return null;
}

function formatDateForTable(date) {
    if (!date || isNaN(date.getTime())) return 'N/A';
    try {
        const year = date.getUTCFullYear();
        const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
        const day = date.getUTCDate().toString().padStart(2, '0');
        return `${year}-${month}-${day}`;
    } catch (e) {
        Logger.log(`Error formatting date ${date}: ${e}`);
        return 'N/A';
    }
}

function getCurrentDateInfo() {
  const today = new Date();
  const dateString = today.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const d = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return { date: dateString, week: `Week ${weekNo}` };
}

function exportToCsv(dataJSON, keysJSON, titlesJSON) {
  try {
    const data = JSON.parse(String(dataJSON));
    const keys = JSON.parse(String(keysJSON));
    const titles = JSON.parse(String(titlesJSON));

    const escapeCsv = (val) => {
      if (val === null || val === undefined) return "";
      let str = String(val);
      if (str.includes('"')) {
         str = str.replace(/"/g, '""');
      }
      if (str.includes(",") || str.includes("\n") || str.includes('"')) {
         str = `"${str}"`; 
      }
      return str;
    };

    let csv = titles.map(escapeCsv).join(",") + "\r\n"; 

    data.forEach(row => {
      const line = keys.map(key => {
        let value = key.split('.').reduce((obj, prop) => (obj && obj.hasOwnProperty(prop)) ? obj[prop] : '', row);
        return escapeCsv(value);
      });
      csv += line.join(",") + "\r\n"; 
    });

    const csvBlob = Utilities.newBlob(csv, MimeType.CSV, "export.csv");
    const base64Data = Utilities.base64Encode(csvBlob.getBytes());
   
    if (!base64Data) {
        throw new Error("Base64 encoding failed.");
    }

    return "data:text/csv;charset=utf-8;base64," + base64Data;
  } catch (e) {
    Logger.log(`Error in exportToCsv: ${e.stack}`);
    return JSON.stringify({ error: `CSV Export Failed: ${e.message}` });
  }
}

const PROJECT_ID = "cp-sandbox-devin-lin687";
const AGENT_ID = "bd2c8cb2-051e-4c2a-a425-8a594781c803";
const LOCATION_ID = "us-central1"; 

const CHAT_HISTORY_SHEET_NAME = "GeminiChatHistory";

function getAccessToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const keyString = scriptProperties.getProperty('SERVICE_ACCOUNT_KEY');

  if (!keyString) {
    throw new Error("SERVICE_ACCOUNT_KEY property is missing. Please add it in Project Settings > Script Properties.");
  }

  let serviceAccount;
  try {
    serviceAccount = JSON.parse(keyString);
  } catch (e) {
    throw new Error("Error parsing SERVICE_ACCOUNT_KEY. Please ensure the value in Script Properties starts with '{' and ends with '}'.");
  }

  const privateKey = serviceAccount.private_key;
  const clientEmail = serviceAccount.client_email;
  if (!privateKey || !clientEmail) {
    throw new Error("SERVICE_ACCOUNT_KEY is valid JSON but missing 'private_key' or 'client_email'.");
  }

  const tokenUri = "https://oauth2.googleapis.com/token";
  const jwtHeader = { "alg": "RS256", "typ": "JWT" };
  const now = Math.floor(Date.now() / 1000);
  const jwtClaimSet = {
    "iss": clientEmail,
    "scope": "https://www.googleapis.com/auth/dialogflow",
    "aud": tokenUri,
    "exp": now + 3600,
    "iat": now
  };
  const headerBase64 = base64urlEncode_(JSON.stringify(jwtHeader));
  const claimSetBase64 = base64urlEncode_(JSON.stringify(jwtClaimSet));
  const signatureInput = headerBase64 + "." + claimSetBase64;
  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const signatureBase64 = base64urlEncode_(signature);
  
  const jwt = signatureInput + "." + signatureBase64;
  const options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": {
      "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
      "assertion": jwt
    },
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(tokenUri, options);
  const responseText = response.getContentText();
  
  if (response.getResponseCode() !== 200) {
    console.error("Auth Error Body:", responseText);
    throw new Error("Failed to get Access Token: " + responseText);
  }

  const data = JSON.parse(responseText);
  return data.access_token;
}

function base64urlEncode_(input) {
  let base64String;
  if (Array.isArray(input) || (typeof input === 'object' && input !== null)) {
    base64String = Utilities.base64Encode(input);
  } else {
    base64String = Utilities.base64Encode(input, Utilities.Charset.UTF_8);
  }
  return base64String.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function generateAgentResponse(userMessage, sessionId) {
  if (!sessionId) {
    sessionId = Utilities.getUuid();
  }

  const url = `https://${LOCATION_ID}-dialogflow.googleapis.com/v3/projects/${PROJECT_ID}/locations/${LOCATION_ID}/agents/${AGENT_ID}/sessions/${sessionId}:detectIntent`;
  try {
    const accessToken = getAccessToken();

    const payload = {
      "queryInput": {
        "text": {
          "text": userMessage
        },
        "languageCode": "en"
      }
    };
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'headers': {
        'Authorization': 'Bearer ' + accessToken 
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true 
    };
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        console.error("Vertex AI API HTTP Error Response:", response.getResponseCode(), response.getContentText());
        if (responseData && responseData.error && responseData.error.message) {
            return `Error communicating with the conversational agent: ${responseData.error.message}`;
        }
        return `Unknown HTTP error communicating with the conversational agent: ${response.getResponseCode()}`;
    }

    if (responseData.queryResult && responseData.queryResult.responseMessages) {
      if (responseData.queryResult.responseMessages.length > 0) {
        const firstMessage = responseData.queryResult.responseMessages[0];
        if (firstMessage.text && firstMessage.text.text && firstMessage.text.text.length > 0) {
          return firstMessage.text.text.join('\n');
        } else if (firstMessage.payload) {
          console.log("Agent returned a custom payload:", firstMessage.payload);
          return "Agent returned custom information, please check the console.";
        } else {
          console.warn("Agent returned a response, but the first message does not contain text or custom payload:", firstMessage);
          return "Agent returned a response, but its format does not contain recognizable text.";
        }
      } else {
        console.warn("Agent returned an empty responseMessages array:", responseData);
        return "Agent returned an empty response.";
      }
    } else if (responseData.error) {
      console.error("Vertex AI API Error:", responseData.error.message, responseData.error.details);
      if (responseData.error.details && responseData.error.details.length > 0) {
        const errorDetail = responseData.error.details[0];
        if (errorDetail.message) {
          return `Error communicating with the conversational agent: ${errorDetail.message}`;
        }
      }
      return `Error communicating with the conversational agent: ${responseData.error.message || 'Unknown error'}`;
    } else {
      console.error("Unknown API response structure:", responseData);
      return "Sorry, received an unexpected response structure from the agent.";
    }
  } catch (e) {
    console.error("Error calling Vertex AI API:", e.toString());
    return "An error occurred while connecting to the AI, please try again. Error details: " + e.message;
  }
}

function getUserId() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail) {
      return userEmail;
    }
  } catch (e) {
    console.log("Could not get user email, possibly a permissions issue. Using random ID.");
  }
  return "anon-" + Utilities.getUuid();
}

function getChatHistorySheet() {
  try {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CHAT_HISTORY_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CHAT_HISTORY_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'User ID', 'Chat Session ID', 'Chat History (JSON)']);
    }
    return sheet;
  } catch (e) {
    const files = DriveApp.getFilesByName(CHAT_HISTORY_SHEET_NAME);
    let spreadsheet;
    if (files.hasNext()) {
      spreadsheet = SpreadsheetApp.open(files.next());
    } else {
      spreadsheet = SpreadsheetApp.create(CHAT_HISTORY_SHEET_NAME);
    }
    let sheet = spreadsheet.getSheetByName(CHAT_HISTORY_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CHAT_HISTORY_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'User ID', 'Chat Session ID', 'Chat History (JSON)']);
    }
    return sheet;
  }
}

function saveChatSession(chatSessionId, userId, chatHistoryJson) {
  const sheet = getChatHistorySheet();
  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId && data[i][2] === chatSessionId) {
      foundRow = i + 1;
      break;
    }
  }
  const rowData = [new Date(), userId, chatSessionId, chatHistoryJson];
  if (foundRow !== -1) {
    sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function loadChatSession(userId) {
  const sheet = getChatHistorySheet();
  const data = sheet.getDataRange().getValues();
  let latestSession = null;
  let latestTimestamp = new Date(0); 

  for (let i = data.length - 1; i > 0; i--) { 
    const row = data[i];
    const rowUserId = row[1]; 
    if (rowUserId === userId) {
      const timestamp = new Date(row[0]);
      if (timestamp > latestTimestamp) {
        latestTimestamp = timestamp;
        latestSession = { sessionId: row[2], history: row[3] }; 
      }
    }
  }
  return latestSession;
}
