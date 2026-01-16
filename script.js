// --- Wait for the page to be fully loaded ---
document.addEventListener("DOMContentLoaded", () => {
  // Get references to all THREE list elements
  // --- Upload Modal Elements ---
  const openUploadBtn = document.getElementById("openUploadBtn");
  const uploadModal = document.getElementById("uploadModal");
  const closeUploadSpan = document.querySelector(".close-upload");
  const fileRowContainer = document.getElementById("fileRowContainer");
  const addFileRowBtn = document.getElementById("addFileRowBtn");
  const processFilesBtn = document.getElementById("processFilesBtn");
  const criticalList = document.getElementById("criticalList");
  const belowList = document.getElementById("belowList");
  const aboveList = document.getElementById("aboveList");
  const lowList = document.getElementById("lowList");
  const moderateList = document.getElementById("moderateList");
  const goodList = document.getElementById("goodList");
  const modal = document.getElementById("reportModal");
  // const modalContent = document.getElementById("reportContent"); // We don't use this ID anymore
  const closeModalButton = document.getElementById("closeModalButton");
  const mainTitle = document.getElementById("mainTitle");

  const reportCountryNameSpan = document.getElementById("reportCountryName");
  const reportDetailsContainer = document.getElementById("reportDetails");
  const countrySearchInput = document.getElementById("countrySearch");

  // --- New Elements for Compare Mode ---
  const compareToggle = document.getElementById("compareToggle");
  const comparisonModal = document.getElementById("comparisonModal");
  const closeCompareModalBtn = document.getElementById("closeCompareModalBtn");
  const compareContainer = document.getElementById("compareContainer");
  // --- Global Trends Elements ---
  // const globalKPISelect = document.getElementById("globalTrendKPISelect");
  const globalCountrySearch = document.getElementById(
    "globalTrendCountrySearch"
  );
  const globalCountryList = document.getElementById("globalTrendCountryList");
  const globalDateRange = document.getElementById("globalTrendDateRange");
  const customDateContainer = document.getElementById("customDateContainer");
  const customDateStart = document.getElementById("customDateStart");
  const customDateEnd = document.getElementById("customDateEnd");
  // --- End New Elements ---

  // --- NEW: State Management ---
  const DASHBOARD_STATE = {}; // Stores data for all uploaded games: { "DIQ-2": { ...data... } }
  let activeGameKey = null; // Tracks which tab is currently open

  let workbook = null;
  let chartInstances = new Map();

  // --- Data for Single Modal ---
  let currentCountryData = null;
  let currentLatestKPIs = null;
  let currentBenchmarkKPIs = null;
  // --- End Data ---

  // --- State Variables ---
  let isCompareMode = false;
  let allCountriesData = new Map(); // Stores all country data for search
  let selectedCountriesForCompare = []; // Array to hold selected country names
  let selectedTrendCountries = new Set(); // For Global Trend Chart

  // --- UPLOAD MODAL LOGIC ---

  // 1. Open Modal
  openUploadBtn.addEventListener("click", () => {
    uploadModal.style.display = "flex";
    uploadModal.classList.remove("modal-hidden");

    // FORCE ADD ROW if empty
    if (fileRowContainer.children.length === 0) {
      addFileRow();
    }
  });

  // 2. Close Modal
  closeUploadSpan.addEventListener("click", () => {
    uploadModal.style.display = "none";
    uploadModal.classList.add("modal-hidden");
  });

  // 3. Add New Row Function (Fixed for all environments)
  addFileRowBtn.addEventListener("click", () => addFileRow());

  // Helper: Checks if ANY file exists (Robust Version)
  function checkButtonState() {
    const rows = fileRowContainer.querySelectorAll(".file-row");
    let hasFiles = false;

    rows.forEach((row) => {
      // Check 1: Our custom property (Multi-select)
      if (row.customFile) hasFiles = true;
      // Check 2: The native input (Single select fallback)
      const input = row.querySelector("input");
      if (input && input.files && input.files.length > 0) hasFiles = true;
    });

    if (hasFiles) {
      processFilesBtn.classList.add("btn-success");
      // Optional: processFilesBtn.textContent = "Process Files ✅";
    } else {
      processFilesBtn.classList.remove("btn-success");
    }
  }

  // Updated: Clean UI (Hides Browse on select), Handles Duplicates
  function addFileRow(autoFile = null) {
    const rowCount = fileRowContainer.children.length + 1;
    const uniqueId =
      "file-" + Date.now() + "-" + Math.floor(Math.random() * 10000);

    const div = document.createElement("div");
    div.className = "file-row";

    if (autoFile) div.customFile = autoFile;

    // Ultra-short text to fit in small boxes
    const placeholderText = "Select File (Multiple OK)";

    div.innerHTML = `
      <span class="row-number">${rowCount}.</span>
      <input type="file" id="${uniqueId}" class="game-file-input" accept=".xlsx, .xls" multiple>
      <label for="${uniqueId}" class="file-label">Browse...</label>
      <span class="file-name-display">${
        autoFile ? autoFile.name : placeholderText
      }</span>
      <button class="remove-row-btn" title="Remove">&times;</button>
    `;

    const nameDisplay = div.querySelector(".file-name-display");
    const fileInput = div.querySelector(".game-file-input");
    const fileLabel = div.querySelector(".file-label"); // Reference to the button
    const removeBtn = div.querySelector(".remove-row-btn");

    // --- VIEW LOGIC: Hide Browse button & Box the Text ---
    const updateViewState = (hasFile) => {
      if (hasFile) {
        fileLabel.style.display = "none"; // Hide the button
        nameDisplay.classList.add("file-filled"); // Add the "Box" style
        // Reset manual styles just in case
        nameDisplay.style.color = "";
        nameDisplay.style.fontStyle = "";
      } else {
        fileLabel.style.display = "block"; // Show the button
        nameDisplay.classList.remove("file-filled"); // Remove "Box" style
        // Apply placeholder styles
        nameDisplay.style.color = "#999";
        nameDisplay.style.fontStyle = "italic";
      }
    };

    // Initial State Check
    if (autoFile) {
      updateViewState(true);
      checkButtonState(); // Update button color
    } else {
      updateViewState(false);
    }

    checkButtonState(); // Check button color on init

    fileInput.addEventListener("change", (e) => {
      const rawFiles = Array.from(e.target.files);

      // 1. Gather existing files to check duplicates
      const existingNames = new Set(
        Array.from(fileRowContainer.querySelectorAll(".file-row"))
          .filter((row) => row.customFile)
          .map((row) => row.customFile.name)
      );

      // 2. Filter duplicates
      const validFiles = rawFiles.filter((file) => {
        if (existingNames.has(file.name)) {
          console.log(`Duplicate ignored: ${file.name}`);
          return false;
        }
        return true;
      });

      if (validFiles.length === 0 && rawFiles.length > 0) {
        alert("Skipped duplicate files.");
        fileInput.value = "";
        checkButtonState(); // Re-check
        return;
      }

      if (validFiles.length > 0) {
        // A. Fill THIS row
        const firstFile = validFiles[0];
        div.customFile = firstFile;
        nameDisplay.textContent = firstFile.name;

        // Update UI to clean mode
        updateViewState(true);

        // FORCE BUTTON CHECK NOW
        setTimeout(checkButtonState, 0);

        // B. Spawn new rows for extra files
        if (validFiles.length > 1) {
          const remainingFiles = validFiles.slice(1);
          remainingFiles.forEach((file) => addFileRow(file));
        }
      } else {
        // C. User cancelled
        nameDisplay.textContent = placeholderText;
        div.customFile = null;
        updateViewState(false);
        checkButtonState(); // Re-check
      }
    });

    removeBtn.addEventListener("click", () => {
      div.remove();
      updateRowNumbers();
      checkButtonState();
    });

    fileRowContainer.appendChild(div);
  }

  // Helper to re-number rows nicely
  function updateRowNumbers() {
    const rows = fileRowContainer.querySelectorAll(".file-row");
    rows.forEach((row, index) => {
      row.querySelector(".row-number").textContent = index + 1 + ".";
    });
  }

  // --- INSERT THIS BLOCK AT LINE 108 (After addFileRow function) ---

  // 4. Process Files Logic (Updated: Clears old data on every run)
  processFilesBtn.addEventListener("click", () => {
    const rows = fileRowContainer.querySelectorAll(".file-row");

    // Step 1: Check if there are actually files to process
    let filesFound = false;
    rows.forEach((row) => {
      if (row.querySelector(".game-file-input").files.length > 0)
        filesFound = true;
    });

    if (!filesFound) {
      alert("Please select at least one file to process.");
      return;
    }

    // Step 2: CLEAR OLD DATA (The Fix)
    // We empty the dashboard state so deleted games disappear
    Object.keys(DASHBOARD_STATE).forEach((key) => delete DASHBOARD_STATE[key]);
    activeGameKey = null;
    workbook = null; // Clear global workbook reference
    allCountriesData.clear(); // Clear search map
    renderTabs();

    // Step 3: Process the current files
    rows.forEach((row) => {
      const fileInput = row.querySelector(".game-file-input");

      // PRIORITY: Check if we have a stored file (from multi-select spillover)
      // Fallback: Check the input field (standard single select)
      let fileToProcess = row.customFile;

      if (!fileToProcess && fileInput.files.length > 0) {
        fileToProcess = fileInput.files[0];
      }

      if (fileToProcess) {
        const reader = new FileReader();

        reader.onload = (e) => {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });

          // Use the name from the file object
          processBenchmarkSheet(fileToProcess.name, wb);
        };

        reader.readAsArrayBuffer(fileToProcess);
      }
    });

    // Step 4: Close modal
    uploadModal.style.display = "none";
    uploadModal.classList.add("modal-hidden");
  });

  // --- END INSERTION ---
  // --- Listen for Compare Toggle ---
  compareToggle.addEventListener("change", (event) => {
    isCompareMode = event.target.checked;
    if (isCompareMode) {
      // Clear filter when enabling compare mode
      countrySearchInput.value = "";
      filterCountries(); // Resets the list
    } else {
      countrySearchInput.placeholder = "Search for a country...";
      // --- START: Added Reset Logic ---
      clearCompareSelections();
      // --- END: Added Reset Logic ---
    }
  });

  countrySearchInput.addEventListener("input", () => {
    // Filter regardless of compare mode
    filterCountries();
  });

  // Helper to convert ISO dates (2026-01-04) to Sheet Name format (4 Jan 26)
  function normalizeDateForSheet(val) {
    if (!val) return "";
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    const day = d.getDate();
    const months = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];
    const month = months[d.getMonth()];
    const year = String(d.getFullYear()).slice(-2);
    return `${day} ${month} ${year}`;
  }
  // --- NEW: Render Tabs ---
  function renderTabs() {
    const container = document.getElementById("tabsContainer");
    container.innerHTML = ""; // Clear existing tabs

    Object.keys(DASHBOARD_STATE)
      .sort()
      .forEach((key) => {
        const btn = document.createElement("button");
        btn.className = `tab-button ${key === activeGameKey ? "active" : ""}`;
        btn.textContent = key; // e.g., "DIQ-2"
        btn.onclick = () => switchTab(key);
        container.appendChild(btn);
      });
  }

  // --- NEW: Switch Game Context ---
  function switchTab(key) {
    if (!DASHBOARD_STATE[key]) return;

    activeGameKey = key;
    const data = DASHBOARD_STATE[key];

    // 1. Restore Global State for this game
    // We update the helper map so search works for THIS game
    allCountriesData = data.allCountriesMap;
    // IMPORTANT: If your other functions rely on 'workbook', you might need to pass it
    // or set a temporary global. For now, we update the module-level 'workbook' variable if you kept it,
    // OR update your helper functions to accept 'workbook'.
    // *Simplest Fix:* Update the global 'workbook' variable to match this tab:
    workbook = data.workbook;

    // 2. Update UI
    mainTitle.innerHTML = `${data.title} <span style="font-size: 0.6em; color: #737373; vertical-align: middle;">(Last ${data.totalDays} Days)</span>`;
    document.title = `AAPU Report | ${key}`;

    // 3. Render Lists
    displayCountryLists(data.critical, data.below, data.above);
    displayAAPULists(data.low, data.moderate, data.good);

    // 4. Update Tab Styles
    renderTabs(); // Re-runs to highlight the active button

    // 5. Reset Search & Compare
    const searchInput = document.getElementById("countrySearch");
    searchInput.value = "";
    compareToggle.checked = false;
    isCompareMode = false;
    clearCompareSelections(); // Ensure this function exists from your previous code

    // Filter to show all (since we cleared search)
    filterCountries();

    // SMART RETENTION: Keep current selections if they exist in the new game data
    const newSelected = new Set();
    selectedTrendCountries.forEach((nameKey) => {
      if (allCountriesData.has(nameKey)) {
        // Check if it meets the >10 installs rule in the NEW game context
        if (allCountriesData.get(nameKey).userInstalls > 10) {
          newSelected.add(nameKey);
        }
      }
    });
    selectedTrendCountries = newSelected;

    if (globalCountrySearch) globalCountrySearch.value = "";
    populateTrendCountryList();
    updateGlobalTrendChart();
  }

  // --- Function to process the 'BM' sheet (FIXED: Silent Mode) ---
  function processBenchmarkSheet(fileName, wb) {
    const bmSheetName = wb.SheetNames.find(
      (name) => name.toUpperCase() === "BM"
    );
    if (!bmSheetName) {
      alert(
        `Error: The file "${fileName}" does not contain a "BM" sheet.\n\nPlease upload the processed "All Geo.xlsx" file, not the raw CSV.`
      );
      return;
    }

    const ws = wb.Sheets[bmSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const headers = data[0].map((h) => {
      // Convert Excel serial numbers (like 46026) to JS Date strings
      if (typeof h === "number" && h > 40000) {
        return new Date(Math.round((h - 25569) * 86400 * 1000))
          .toISOString()
          .split("T")[0];
      }
      return h;
    });
    const latestDate = headers[1];
    const latestDateCol = 1;
    const totalDays = headers.length - 1;

    // --- Determine Game Name from Filename ---
    let gameName = "";

    if (fileName) {
      const upperName = fileName.toUpperCase();
      if (upperName.includes("DIQ-1")) {
        gameName = "DIQ-1";
      } else if (upperName.includes("IOS")) {
        gameName = "DIQ-2 iOS";
      } else if (upperName.includes("DIQ-2")) {
        gameName = "DIQ-2";
      } else if (upperName.includes("DIQ-3")) {
        gameName = "DIQ-3";
      }
    }

    // --- FIX 1: Create LOCAL Map (Don't use the global one yet) ---
    const localCountriesMap = new Map();

    const criticalCountries = [];
    const belowCountries = [];
    const aboveCountries = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0 || !row[0]) continue;

      let country = row[0];

      try {
        const latestValue = row[latestDateCol];

        const normalizedLatest = normalizeDateForSheet(latestDate);
        const latestSheetName = wb.SheetNames.find((name) =>
          name.toLowerCase().includes(normalizedLatest.toLowerCase())
        );

        if (!latestSheetName) continue;

        let userInstalls = 0;
        let onboardUsers = 0;
        if (latestSheetName) {
          const latestKPIs = getKPIsForCountry(latestSheetName, country, wb);
          if (latestKPIs) {
            userInstalls = Number(latestKPIs["User Installed"]) || 0;
            onboardUsers = Number(latestKPIs["Users Onboarded"]) || 0;
          }
        }

        if (userInstalls <= 10) continue;

        let benchmarkValue = -1.0;
        let benchmarkDate = "N_A";

        for (let j = 2; j < headers.length; j++) {
          const currentValue = row[j];
          if (
            typeof currentValue === "number" &&
            currentValue > benchmarkValue
          ) {
            benchmarkValue = currentValue;
            benchmarkDate = headers[j];
          }
        }

        let isNewBenchmark = false;
        if (typeof latestValue === "number" && latestValue >= benchmarkValue) {
          isNewBenchmark = true;
        }

        const trendDates = headers.slice(1, 8).reverse();

        const trendValuesRaw = row
          .slice(1, 8)
          .map((v) => (typeof v === "number" ? v : null))
          .reverse();

        const trendValuesFormatted = trendValuesRaw.map((v) =>
          v !== null ? v.toFixed(2) : null
        );

        if (typeof latestValue === "number" && benchmarkValue > -1) {
          const diff = benchmarkValue - latestValue;

          const countryData = {
            name: country,
            latestAAPU: latestValue,
            benchmarkAAPU: benchmarkValue,
            latestDate: latestDate,
            benchmarkDate: benchmarkDate,
            isCritical: diff >= 1.0,
            trendDates: trendDates,
            trendValuesRaw: trendValuesRaw,
            trendValuesFormatted: trendValuesFormatted,
            totalDays: totalDays,
            isNewBenchmark: isNewBenchmark,
            userInstalls: userInstalls,
            onboardUsers: onboardUsers,
          };

          // --- FIX 2: Use Local Map ---
          localCountriesMap.set(country.toLowerCase(), countryData);

          if (latestValue < benchmarkValue) {
            if (countryData.isCritical) {
              criticalCountries.push(countryData);
            } else {
              belowCountries.push(countryData);
            }
          } else {
            aboveCountries.push(countryData);
          }
        }
      } catch (error) {
        console.error(`Error processing country: ${country}`, error);
      }
    }

    // --- FIX 3: REMOVED ALL DIRECT DOM UPDATES (displayCountryLists, etc.) ---

    const low = [],
      moderate = [],
      good = [];
    localCountriesMap.forEach((c) => {
      if (c.latestAAPU < 7) low.push(c);
      else if (c.latestAAPU <= 8) moderate.push(c);
      else good.push(c);
    });

    const key = gameName || fileName;

    DASHBOARD_STATE[key] = {
      workbook: wb,
      title: `${gameName} AAPU Performance Dashboard of ${normalizeDateForSheet(
        latestDate
      )}`,
      totalDays: totalDays,
      critical: criticalCountries,
      below: belowCountries,
      above: aboveCountries,
      low: low,
      moderate: moderate,
      good: good,
      allCountriesMap: localCountriesMap,
      history: crawlHistory(wb), // Build full historical map for all KPIs
    };

    renderTabs();

    // Only switch if this is the FIRST game loaded (avoids hijacking)
    if (!activeGameKey) {
      switchTab(key);
    }
  }

  // --- Function to display THREE lists ---
  function displayCountryLists(critical, below, up) {
    criticalList.innerHTML = "";
    belowList.innerHTML = "";
    aboveList.innerHTML = "";

    // Helper function to create list items
    const createListItem = (country, listElement, isAbove = false) => {
      const item = document.createElement("div");
      item.className = `country-item ${isAbove ? "above" : ""}`; // Add 'above' class if needed

      // *** REPLACE START ***
      // Decide what benchmark text to show
      let bmText = "";
      if (isAbove && country.isNewBenchmark) {
        // If it's in the 'Above' list AND it set a new benchmark today
        bmText = `🚀 New BM!`; // Show latest value as the new BM
      } else {
        // Otherwise (it's below, or it's above but didn't set a new BM today)
        bmText = `(BM: ${country.benchmarkAAPU.toFixed(2)})`; // Show the calculated historical BM
      }

      // Assign the final HTML
      item.innerHTML = `
          <strong>${country.name}</strong>
          <span>${country.latestAAPU.toFixed(
            2
          )} <small>${bmText}</small></span> 
 `;
      // *** REPLACE END ***

      // --- START: Modified Click Listener ---
      item.addEventListener("click", () => {
        if (isCompareMode) {
          handleCompareSelection(country, item); // Call new function for compare
        } else {
          showWTHReport(country); // Original behavior
        }
      });
      // --- END: Modified Click Listener ---
      listElement.appendChild(item);
    };

    if (critical.length === 0) {
      criticalList.innerHTML =
        '<p class="placeholder">No critical countries.</p>';
    } else {
      critical.forEach((country) => createListItem(country, criticalList));
    }

    if (below.length === 0) {
      belowList.innerHTML =
        '<p class="placeholder">No countries are slightly below benchmark.</p>';
    } else {
      below.forEach((country) => createListItem(country, belowList));
    }

    if (up.length === 0) {
      aboveList.innerHTML =
        '<p class="placeholder">No countries are above benchmark.</p>';
    } else {
      up.forEach((country) => createListItem(country, aboveList, true)); // Pass true for isAbove
    }
  }

  // --- New Function to filter countries ---
  function filterCountries() {
    const searchTerm = countrySearchInput.value.toLowerCase();

    // An array of the 3 list elements
    const lists = [
      criticalList,
      belowList,
      aboveList,
      lowList,
      moderateList,
      goodList,
    ];

    lists.forEach((listElement) => {
      const allItems = listElement.querySelectorAll(".country-item");
      let visibleCount = 0;

      // If the list has items, filter them
      if (allItems.length > 0) {
        allItems.forEach((item) => {
          // --- START: Added Check for Selected Item ---
          // If in compare mode AND this item is already selected, ALWAYS show it
          if (
            isCompareMode &&
            item.classList.contains("selected-for-compare")
          ) {
            item.style.display = "flex";
            visibleCount++; // Ensure placeholder logic still works
            return; // Skip the rest of the filtering for this item
          }
          // --- END: Added Check ---

          const countryName = item
            .querySelector("strong")
            .textContent.toLowerCase();
          if (countryName.includes(searchTerm)) {
            item.style.display = "flex";
            visibleCount++;
          } else {
            item.style.display = "none";
          }
        });

        // Now, manage the "No results" placeholder
        let placeholder = listElement.querySelector(".placeholder");
        if (visibleCount === 0) {
          // All items are hidden, so show a "no results" message
          if (!placeholder) {
            // Create a placeholder if one doesn't exist
            placeholder = document.createElement("p");
            placeholder.className = "placeholder";
            listElement.appendChild(placeholder);
          }
          placeholder.textContent = "No matching countries found.";
          placeholder.style.display = "block";
        } else if (placeholder) {
          // Items are visible, hide any placeholder
          placeholder.style.display = "none";
        }
      }
      // If allItems.length is 0, the list was empty to begin with,
      // so we let the original placeholder ("No critical countries.") remain.
    });
  }

  // --- NEW: Handle clicks in Compare Mode ---
  function handleCompareSelection(country, itemElement) {
    const countryName = country.name; // Use the actual name
    const index = selectedCountriesForCompare.indexOf(countryName);

    if (index > -1) {
      // Already selected - Deselect it
      selectedCountriesForCompare.splice(index, 1);
      itemElement.classList.remove("selected-for-compare");
    } else {
      // Not selected - Try to select it
      if (selectedCountriesForCompare.length < 2) {
        selectedCountriesForCompare.push(countryName);
        itemElement.classList.add("selected-for-compare");

        // If two selected, wait briefly then show report and reset
        if (selectedCountriesForCompare.length === 2) {
          // --- START: Added Delay ---
          setTimeout(() => {
            // Pass the names (lowercase) to the existing function
            showComparisonReport(
              selectedCountriesForCompare.map((name) => name.toLowerCase())
            );
            // Reset after showing report
            clearCompareSelections();
          }, 300); // Wait 500ms (half a second)
          // --- END: Added Delay ---
        }
      } else {
        // Already 2 selected
        alert("You can only compare two countries at a time.");
      }
    }
  }
  // --- NEW: Helper to clear selections ---
  function clearCompareSelections() {
    selectedCountriesForCompare = []; // Reset array
    // Remove visual style from all items
    const allSelectedItems = document.querySelectorAll(
      ".country-item.selected-for-compare"
    );
    allSelectedItems.forEach((item) =>
      item.classList.remove("selected-for-compare")
    );
  }
  function showWTHReport(country) {
    const mainChart = chartInstances.get("main");
    if (mainChart) {
      mainChart.destroy();
      chartInstances.delete("main");
    }

    // --- 1. Find sheets (using normalized year-aware names) ---
    const normLatest = normalizeDateForSheet(country.latestDate);
    const normBench = normalizeDateForSheet(country.benchmarkDate);

    const latestSheetName = workbook.SheetNames.find((name) =>
      name.toLowerCase().includes(normLatest.toLowerCase())
    );
    const benchmarkSheetName = workbook.SheetNames.find((name) =>
      name.toLowerCase().includes(normBench.toLowerCase())
    );

    if (!latestSheetName) {
      reportDetailsContainer.innerHTML = `<p style="color:red;">Error: Could not find data sheet for latest date: "${country.latestDate}"</p>`;
      modal.style.display = "flex";
      modal.classList.remove("modal-hidden");
      return;
    }

    if (!benchmarkSheetName && country.benchmarkDate !== "N_A") {
      reportDetailsContainer.innerHTML = `<p style="color:red;">Error: Could not find data sheet for benchmark date: "${country.benchmarkDate}"</p>`;
      modal.style.display = "flex";
      modal.classList.remove("modal-hidden");
      return;
    }

    // Pass workbook explicitly to avoid tab-sync issues
    const latestKPIs = getKPIsForCountry(
      latestSheetName,
      country.name,
      workbook
    );
    const benchmarkKPIs = benchmarkSheetName
      ? getKPIsForCountry(benchmarkSheetName, country.name, workbook)
      : null;

    // --- 2. Store data globally for toggle ---
    currentCountryData = country;
    currentLatestKPIs = latestKPIs;
    currentBenchmarkKPIs = benchmarkKPIs;

    // --- 3. Set Defaults and Initial Display ---
    //reportCountryNameSpan.textContent = country.name; // Set country name in header

    // *** NEW: Update the static summary in the header ***
    const summaryP = document.querySelector(
      ".dp1-modal-content .report-header p.summary"
    );
    if (summaryP) {
      const isDown =
        !country.isNewBenchmark && country.latestAAPU < country.benchmarkAAPU;
      const statusClass = isDown ? "down" : "up";
      const benchmarkFullText = country.isNewBenchmark
        ? `(🚀 New Benchmark!)`
        : `(Benchmark: ${country.benchmarkAAPU.toFixed(
            2
          )} on ${normalizeDateForSheet(country.benchmarkDate)})`;
      summaryP.innerHTML = `
        <strong>Current AAPU: <span class="${statusClass}">${country.latestAAPU.toFixed(
        2
      )}</span></strong> 
        <small>${benchmarkFullText}</small>
    `;
    }

    // --- 4. Call the *refactored* update function ---
    updateReportContent(
      reportDetailsContainer, // The container to update
      currentCountryData,
      currentLatestKPIs,
      currentBenchmarkKPIs,
      false, // Default to "vs Install"
      "main" // The ID for this report
    );

    // --- 4. Show Modal ---
    modal.style.display = "flex";
    modal.classList.remove("modal-hidden");
  }

  // --- Helper function to get KPIs ---
  function getKPIsForCountry(sheetName, countryName, wbOverride = null) {
    // FIX: Use the override if provided, otherwise use global workbook
    const currentWb = wbOverride || workbook;

    if (!currentWb) return null; // Safety check

    const ws = currentWb.Sheets[sheetName]; // Use currentWb
    if (!ws) return null;

    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (data.length === 0) return null;

    const headers = data[0];
    const kpiColumnNames = data.map((row) => row[0]);

    let countryCol = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === countryName) {
        countryCol = i;
        break;
      }
    }

    if (countryCol === -1) return null;

    const kpis = {};
    kpiColumnNames.forEach((kpiName, rowIndex) => {
      if (kpiName && rowIndex > 0) {
        kpis[kpiName] = data[rowIndex][countryCol];
      }
    });
    return kpis;
  }

  // --- Function to generate the HTML report content ---
  function generateReportHTML(country, latest, bench, vsOnboarded, reportId) {
    if (!latest)
      return `<h3>Error</h3><p>Missing KPI data for ${country.name} on ${country.latestDate}.</p>`;

    // Use historical benchmark for status color (isDown)
    const isDown =
      !country.isNewBenchmark && country.latestAAPU < country.benchmarkAAPU;
    const statusClass = isDown ? "down" : "up";

    // --- 1. Define Keys based on Toggle ---
    // *** CRITICAL: You MUST change "vs Onboard - Lvl XX %" to match your Excel columns EXACTLY ***
    const key_lvl20 = vsOnboarded
      ? "vs Onboard - Lvl 20 %"
      : "% of users at 20";
    const key_lvl50 = vsOnboarded
      ? "vs Onboard - Lvl 50 %"
      : "% of users at 50";
    const key_lvl70 = vsOnboarded
      ? "vs Onboard - Lvl 70 %"
      : "% of users at 70";
    const key_lvl100 = vsOnboarded
      ? "vs Onboard - Lvl 100 %"
      : "% of users at 100";
    const key_lvl150 = vsOnboarded
      ? "vs Onboard - Lvl 150 %"
      : "% of users at 150";
    const key_lvl200 = vsOnboarded
      ? "vs Onboard - Lvl 200 %"
      : "% of users at 200";
    // --- End of Key Definitions ---

    // --- 2. KPI 'What Happened' Section ---
    const kpiMap = [
      { name: "Retention Lvl 20", key: key_lvl20, type: "ret" },
      { name: "Retention Lvl 50", key: key_lvl50, type: "ret" },
      { name: "Retention Lvl 70", key: key_lvl70, type: "ret" },
      { name: "Retention Lvl 100", key: key_lvl100, type: "ret" },
      { name: "Retention Lvl 150", key: key_lvl150, type: "ret" },
      { name: "Retention Lvl 200", key: key_lvl200, type: "ret" },
      // Ads are NOT toggled
      { name: "Ads 10 Watched", key: "% of users at Ads 10", type: "ad" },
      { name: "Ads 20 Watched", key: "% of users at Ads 20", type: "ad" },
      { name: "Ads 40 Watched", key: "% of users at Ads 40", type: "ad" },
      { name: "Ads 70 Watched", key: "% of users at Ads 70", type: "ad" },
      { name: "Ads 100 Watched", key: "% of users at Ads 100", type: "ad" },
    ];

    const retKPIs = kpiMap.filter((k) => k.type === "ret");
    const adKPIs = kpiMap.filter((k) => k.type === "ad");

    // --- Helper function to generate list items ---
    const generateListItems = (kpiList, isRetentionList) => {
      // Regular Day: Show comparison logic
      return kpiList
        .map((k) => {
          const l = latest[k.key] || 0; // This is the latest day's value
          const benchKey = k.key; // The key is the same for latest and bench

          let b = 0;
          let benchValueExists = false;

          // Check if benchmark data AND the specific key (e.g., "vs Onboard - Lvl 20 %") exist
          if (bench && bench[benchKey] !== undefined) {
            b = bench[benchKey] || 0; // This is the benchmark day's value
            benchValueExists = true;
          }

          // Calculate change only if we have a benchmark value
          const change_pp = benchValueExists ? (l - b) * 100 : 0;

          let cssClass = "";
          let kpiWarningSymbol = "";

          // Apply coloring/warnings ONLY if the benchmark value exists
          if (benchValueExists) {
            if (change_pp < -1.0) {
              cssClass = "down";
              if (change_pp < -4.0) {
                kpiWarningSymbol = " ❗";
              }
            } else if (change_pp > 1.0) {
              cssClass = "up";
            }
          }

          // Determine the text to show in the brackets
          const bmLabel = country.isNewBenchmark ? "Old BM" : "BM"; // Check if it's a new BM day

          const bmText = benchValueExists
            ? `(${(b * 100).toFixed(2)}% ${bmLabel})` // Use the new label
            : "(BM N/A)"; // e.g., (BM N/A)

          return `<li>${k.name}: <strong class="${cssClass}">${(
            l * 100
          ).toFixed(
            2
          )}%${kpiWarningSymbol}</strong> <small>${bmText}</small></li>`;
        })
        .join("");
    }; // --- End of generateListItems ---

    // --- Generate Lists ---
    const retListHTML = generateListItems(retKPIs, true); // Pass 'true' as it's the retention list
    const adListHTML = generateListItems(adKPIs, false); // Pass 'false' as it's the ad list

    // Trend summary text removed as per request
    let trendSummary = "";

    // --- 3. *** FINAL Dynamic Analysis Section (Includes New BM check) *** ---
    let analysisTitle = "Analysis";
    let analysisAdvice = "";
    let whatHappenedTitle = isDown ? "What Happened" : "What's Working";

    // *** ADDED: Special Case for New Benchmark ***
    if (country.isNewBenchmark) {
      analysisTitle = "🚀 New Benchmark Achieved!";
      analysisAdvice =
        "<p>Performance hit a new peak today! Analyze the <strong>Levels 1-20 experience</strong> and recent changes to understand what drove this success and how to maintain it.</p>";
      whatHappenedTitle = "Today's Performance (New Benchmark)"; // Adjust title
    }
    // *** WRAPPED original logic in 'else' ***
    else if (bench) {
      // Only run detailed analysis if benchmark data exists AND it's not a new BM day
      // Calculate differences for key metrics
      const diff_lvl20 =
        ((latest["% of users at 20"] || 0) - (bench["% of users at 20"] || 0)) *
        100;
      const diff_lvl50 =
        ((latest["% of users at 50"] || 0) - (bench["% of users at 50"] || 0)) *
        100;
      const diff_lvl70 =
        ((latest["% of users at 70"] || 0) - (bench["% of users at 70"] || 0)) *
        100;
      const diff_ads10 =
        ((latest["% of users at Ads 10"] || 0) -
          (bench["% of users at Ads 10"] || 0)) *
        100;
      const diff_ads20 =
        ((latest["% of users at Ads 20"] || 0) -
          (bench["% of users at Ads 20"] || 0)) *
        100;

      // Flags for SIGNIFICANT drops (>4pp)
      const earlyRetSigDrop = diff_lvl20 < -4.0;
      const midRetSigDrop = diff_lvl50 < -4.0 || diff_lvl70 < -4.0;
      const earlyAdsSigDrop = diff_ads10 < -4.0 || diff_ads20 < -4.0;

      // Flags for ANY drop (>1pp, needed for context)
      const earlyRetSlightDrop = diff_lvl20 < -1.0;
      const midRetSlightDrop = diff_lvl50 < -1.0 || diff_lvl70 < -1.0;
      const earlyAdsSlightDrop = diff_ads10 < -1.0 || diff_ads20 < -1.0;

      if (isDown) {
        // --- AAPU IS DOWN ---

        // Priority 1: Early Retention Failure (>4pp drop)
        if (earlyRetSigDrop) {
          analysisTitle = "How to Improve (❗ Early Retention Failure)";
          analysisAdvice =
            "<p>The primary driver is a <strong>significant (>4%) drop</strong> in users reaching Level 20. This cripples the entire funnel. <strong>Focus all investigation on the Level 1-20 experience</strong>.</p>";

          // Priority 2: Mid Retention Failure (>4pp drop in Lvl 50 or 70)
        } else if (midRetSigDrop) {
          analysisTitle = "How to Improve (❗ Mid-Retention Failure)";
          analysisAdvice =
            "<p>Early retention is stable, but a <strong>significant (>4%) drop</strong> occurs between <strong>Levels 20-70</strong>. This mid-game failure is likely the main reason fewer users reach later ad milestones, hurting AAPU. Investigate difficulty or engagement in this range.</p>";

          // Priority 3: Monetization Problem (Retention stable/minor drop, but Ads show >4pp drop)
        } else if (
          !earlyRetSlightDrop &&
          !midRetSlightDrop &&
          earlyAdsSigDrop
        ) {
          analysisTitle = "How to Improve (❗ Monetization Problem)";
          analysisAdvice = `<p>Retention throughout the early/mid game looks stable (drops <1%), but early ad engagement (Ads 10/20) is <strong>down significantly (>4%)</strong>. This points strongly to an ad delivery or engagement issue, not retention. Check:</p><ol><li><strong>Session Length</strong> vs BM?</li><li><strong>Ad Failures</strong> ('None' inter_show_req) vs BM?</li><li><strong>Playtime</strong> vs BM?</li><li><strong>Ad Placement/Frequency:</strong> Ads showing correctly?</li></ol>`;

          // Case 4: Mild Drops Analysis (Drops are >1pp but none are >4pp)
        } else if (
          earlyRetSlightDrop ||
          midRetSlightDrop ||
          earlyAdsSlightDrop
        ) {
          analysisTitle = "How to Improve (Mild Drops)";

          // Find the largest magnitude mild drop among key metrics
          let maxMildDrop = 0;
          let primaryMildIssue = "retention"; // Default assumption

          if (diff_lvl20 < -1.0 && Math.abs(diff_lvl20) > maxMildDrop) {
            maxMildDrop = Math.abs(diff_lvl20);
            primaryMildIssue = "early_retention";
          }
          if (diff_lvl50 < -1.0 && Math.abs(diff_lvl50) > maxMildDrop) {
            maxMildDrop = Math.abs(diff_lvl50);
            primaryMildIssue = "mid_retention";
          }
          if (diff_lvl70 < -1.0 && Math.abs(diff_lvl70) > maxMildDrop) {
            if (
              primaryMildIssue !== "mid_retention" ||
              Math.abs(diff_lvl70) > Math.abs(diff_lvl50) + 0.5
            ) {
              maxMildDrop = Math.abs(diff_lvl70);
              primaryMildIssue = "mid_retention";
            }
          }
          if (diff_ads10 < -1.0 && Math.abs(diff_ads10) > maxMildDrop) {
            if (Math.abs(diff_ads10) > maxMildDrop + 0.5) {
              maxMildDrop = Math.abs(diff_ads10);
              primaryMildIssue = "ads";
            }
          }
          if (diff_ads20 < -1.0 && Math.abs(diff_ads20) > maxMildDrop) {
            if (Math.abs(diff_ads20) > maxMildDrop + 0.5) {
              maxMildDrop = Math.abs(diff_ads20);
              primaryMildIssue = "ads";
            }
          }

          // Generate advice based on the largest mild drop
          if (primaryMildIssue === "early_retention") {
            analysisAdvice =
              "<p>Multiple mild drops (>1% but <4%), but the largest is in <strong>Early Retention (Lvl 20)</strong>. Start investigation there (Levels 1-20).</p>";
          } else if (primaryMildIssue === "mid_retention") {
            analysisAdvice =
              "<p>Multiple mild drops (>1% but <4%), but the largest is in <strong>Mid Retention (Lvl 50/70)</strong>. Focus investigation on Levels 20-70 engagement.</p>";
          } else if (primaryMildIssue === "ads") {
            analysisAdvice = `<p>Multiple mild drops (>1% but <4%), but the largest is in <strong>Early Ad Engagement (Ads 10/20)</strong>. Prioritize checking monetization factors:</p><ol><li><strong>Session Length</strong> vs BM?</li><li><strong>Ad Failures</strong> ('None' inter_show_req) vs BM?</li><li><strong>Playtime</strong> vs BM?</li></ol>`;
          } else {
            analysisAdvice =
              "<p>Multiple mild drops detected. Review overall funnel performance starting from Level 1-20.</p>";
          }

          // Fallback: No significant or mild drops detected in early/mid funnel
        } else {
          analysisTitle = "How to Improve (Late Funnel / Other Issue)";
          analysisAdvice =
            "<p>Early & Mid retention and early ad metrics seem relatively stable compared to the benchmark (<1% drops), yet AAPU is down. Investigate potential issues later in the funnel: check <strong>Lvl 100+ retention</strong> and engagement with later ads like <strong>Ads 40, 70, and 100</strong>.</p>";
        }
      } else {
        // --- AAPU IS UP ---
        if (earlyRetSigDrop) {
          // Early retention down > 4% even though AAPU is up
          analysisTitle = "❗ Warning: Severe Over-Monetization Risk";
          analysisAdvice =
            "<p>AAPU is up, but early retention (Lvl 20) is <strong>down significantly (>4%)</strong>. This is highly unsustainable and indicates aggressive early monetization is driving users away. <strong>Urgently review early-game ad frequency and placements.</strong></p>";
        } else if (earlyRetSlightDrop) {
          // Early retention down > 1%
          analysisTitle = "Warning: Potential Over-Monetization";
          analysisAdvice =
            "<p>AAPU is up, but early retention is <strong>down moderately (>1%)</strong>. Monitor this closely. While revenue is up now, losing users early can hurt long-term. Consider slightly reducing early ad pressure.</p>";
        } else {
          analysisTitle = "What's Working (Healthy Growth)";
          analysisAdvice =
            "<p><strong>Healthy success!</strong> AAPU is up or stable, and early retention is also stable or improving. Good balance. Analyze the <strong>Levels 1-20 experience</strong> to reinforce positive trends.</p>";
        }
      }
    } else {
      // --- BENCHMARK DATA MISSING ---
      analysisTitle = "Analysis (Benchmark Data Missing)";
      analysisAdvice =
        "<p>Could not load benchmark KPI data. Analysis is limited to current performance and recent trend.</p>";
    }
    // --- END Dynamic Analysis ---

    // --- 5. Assemble the Report Object ---
    // This returns the pieces to updateReportContent
    return {
      retListHTML: retListHTML,
      adListHTML: adListHTML,
      trendSummary: trendSummary,
      trendDaysCount: country.totalDays,
      analysisTitle: analysisTitle,
      analysisAdvice: analysisAdvice,
      whatHappenedTitle: whatHappenedTitle,
    };
  }

  // --- Add these new functions ---

  // Function to update the active state of toggle labels
  function updateToggleLabels(isOn, toggleLabelInstall, toggleLabelOnboard) {
    if (toggleLabelInstall && toggleLabelOnboard) {
      toggleLabelInstall.classList.toggle("active", !isOn); // 'vs Install' is active when NOT on
      toggleLabelOnboard.classList.toggle("active", isOn); // 'vs Onboard' is active when ON
    }
  }

  // --- *** MAJOR REFACTOR: updateReportContent *** ---
  // This function is now generic and reusable
  function updateReportContent(
    containerEl,
    country,
    latest,
    bench,
    useVsOnboard,
    reportId
  ) {
    if (!country || !latest) {
      containerEl.innerHTML =
        "<p style='color:red;'>Error: Report data not found.</p>";
      return;
    }

    // Generate the data object containing HTML parts
    const reportData = generateReportHTML(
      country,
      latest,
      bench,
      useVsOnboard,
      reportId // Pass the ID
    );

    const userCountHTML = useVsOnboard
      ? `<small style="font-weight: normal; font-size: 0.7em; color: #666;"> (onboard: ${country.onboardUsers})</small>`
      : `<small style="font-weight: normal; font-size: 0.7em; color: #666;"> (install: ${country.userInstalls})</small>`;

    if (reportId === "main") {
      const nameSpan = document.getElementById("reportCountryName");
      if (nameSpan) nameSpan.innerHTML = `${country.name}${userCountHTML}`;
    }

    // Assemble the final HTML for the details section
    // We build the string in parts, starting with the header
    // Assemble the final HTML for the details section
    let finalDetailsHTML = "";

    // The 'main' reportId uses the modal header, others build their own
    if (reportId !== "main") {
      const isDown =
        !country.isNewBenchmark && country.latestAAPU < country.benchmarkAAPU;
      const statusClass = isDown ? "down" : "up";
      const benchmarkFullText = country.isNewBenchmark
        ? `(🚀 New Benchmark!)`
        : `(Benchmark: ${country.benchmarkAAPU.toFixed(
            2
          )} on ${normalizeDateForSheet(country.benchmarkDate)})`;

      finalDetailsHTML += `
        <div class="report-header">
            <h2>Country Performance: ${country.name}${userCountHTML}</h2>
            <p class="summary">
              <strong>Current AAPU: <span class="${statusClass}">${country.latestAAPU.toFixed(
        2
      )}</span></strong> 
              <small>${benchmarkFullText}</small>
            </p>
        </div>
      `;
    }

    // --- Continue with the rest of the report HTML ---
    finalDetailsHTML += `
        <div class="wth-report">
            <div class="report-subheader">
                <h3>${reportData.whatHappenedTitle}</h3>
                <div class="toggle-area">
                    <span class="toggle-label-small" id="toggleLabelInstall-${reportId}">vs Install</span>
                    <label class="switch switch-small">
                        <input type="checkbox" id="retentionToggle-${reportId}">
                        <span class="slider round"></span>
                    </label>
                    <span class="toggle-label-small" id="toggleLabelOnboard-${reportId}">vs Onboard</span> 
                </div>
            </div>
            <div class="kpi-columns">
                <div class="kpi-column">
                    <h4>Retention KPIs</h4>
                    <ul id="retentionListUL-${reportId}">${reportData.retListHTML}</ul>
                </div>
    `;

    // --- CONDITIONAL BLOCK 1: Ad KPIs ---
    // Only add the "Ad KPIs" column if the toggle is OFF (useVsOnboard is false)
    if (!useVsOnboard) {
      finalDetailsHTML += `
                <div class="kpi-column">
                    <h4>Ad KPIs</h4>
                    <ul>${reportData.adListHTML}</ul>
                </div>
        `;
    }

    finalDetailsHTML += `
            </div> 
    `; // Close .kpi-columns

    // --- CONDITIONAL BLOCK 2: Trend & Analysis ---
    // Only add these sections if the toggle is OFF (useVsOnboard is false)
    if (!useVsOnboard) {
      finalDetailsHTML += `
            <h3>Trend Analysis (last 7 days)</h3> 
            ${reportData.trendSummary} 
            <canvas id="trendChart-${reportId}" height="150"></canvas> 
            `;
      // <-- REMOVED Analysis Title h3
      // <-- REMOVED Analysis Advice
    }

    finalDetailsHTML += `
        </div>
    `; // Close .wth-report

    // Target the correct container
    containerEl.innerHTML = finalDetailsHTML;

    // --- 4. Add Event Listeners for *this* instance ---
    const retentionToggle = document.getElementById(
      `retentionToggle-${reportId}`
    );
    const toggleLabelInstall = document.getElementById(
      `toggleLabelInstall-${reportId}`
    );
    const toggleLabelOnboard = document.getElementById(
      `toggleLabelOnboard-${reportId}`
    );

    if (retentionToggle) {
      retentionToggle.addEventListener("change", (event) => {
        const isChecked = event.target.checked; // true if ON (vs Onboard)

        updateToggleLabels(isChecked, toggleLabelInstall, toggleLabelOnboard);

        // --- START: Smooth Transition Logic ---
        // Use containerEl (the passed-in element)
        containerEl.classList.add("fading-out");

        setTimeout(() => {
          // --- RECURSIVE CALL ---
          // Call itself to only update this specific report instance
          updateReportContent(
            containerEl,
            country,
            latest,
            bench,
            isChecked,
            reportId
          );
          // --- END RECURSIVE CALL ---

          containerEl.classList.remove("fading-out");
        }, 250); // 250ms delay for the fade
      });
    }

    // Set the initial state of the toggle
    if (retentionToggle) {
      retentionToggle.checked = useVsOnboard;
    }
    updateToggleLabels(useVsOnboard, toggleLabelInstall, toggleLabelOnboard);

    /// Set the initial state of the toggle
    if (retentionToggle) {
      retentionToggle.checked = useVsOnboard;
    }
    updateToggleLabels(useVsOnboard, toggleLabelInstall, toggleLabelOnboard);

    // --- 5. Draw Chart for *this* instance ---
    // We only draw the chart if the toggle is OFF
    if (!useVsOnboard) {
      setTimeout(() => {
        const trendChartCanvas = document.getElementById(
          `trendChart-${reportId}` // <-- USES THE DYNAMIC ID
        );
        if (!trendChartCanvas) {
          return; // This is fine, it means we are in toggle mode
        }

        // Destroy existing chart for this ID if it exists
        const oldChart = chartInstances.get(reportId); // <-- USES NEW CHART MAP
        if (oldChart) {
          oldChart.destroy();
          chartInstances.delete(reportId);
        }

        const ctx = trendChartCanvas.getContext("2d");
        const newChart = new Chart(ctx, {
          // <-- CREATES newChart
          type: "line",
          data: {
            // Format each date in the trend array to "4 Jan 26" format
            labels: country.trendDates.map((d) => normalizeDateForSheet(d)),
            datasets: [
              {
                label: `Last ${country.trendDates.length} Days AAPU`,
                data: country.trendValuesFormatted, // <-- USES LOCAL country VARIABLE
                borderColor: "rgba(75, 192, 192, 1)",
                backgroundColor: "rgba(75, 192, 192, 0.2)",
                fill: false,
                tension: 0.1,
                spanGaps: true,
              },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: { mode: "index", intersect: false },
            },
            scales: { y: { beginAtZero: false } },
          },
        });
        // Store the new chart instance
        chartInstances.set(reportId, newChart); // <-- SAVES TO NEW CHART MAP
      }, 150); // Small 50ms delay for chart
    }
  }

  function closeModal() {
    modal.style.display = "none";
    modal.classList.add("modal-hidden");
    const mainChart = chartInstances.get("main");
    if (mainChart) {
      mainChart.destroy();
      chartInstances.delete("main");
    }
  }

  closeModalButton.addEventListener("click", closeModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) {
      closeModal();
    }
  });

  // --- NEW: Close Compare Modal ---
  closeCompareModalBtn.addEventListener("click", () => {
    comparisonModal.style.display = "none";
    comparisonModal.classList.add("modal-hidden");

    // Destroy compare charts
    ["compare1", "compare2", "compare3"].forEach((id) => {
      const chart = chartInstances.get(id);
      if (chart) {
        chart.destroy();
        chartInstances.delete(id);
      }
    });
    compareContainer.innerHTML = ""; // Clear content
  });

  // --- ADD THIS EVENT LISTENER ---
  comparisonModal.addEventListener("click", (event) => {
    if (event.target === comparisonModal) {
      // Check if click is on the background
      // Reuse the logic from the close button listener
      comparisonModal.style.display = "none";
      comparisonModal.classList.add("modal-hidden");

      ["compare1", "compare2", "compare3"].forEach((id) => {
        const chart = chartInstances.get(id);
        if (chart) {
          chart.destroy();
          chartInstances.delete(id);
        }
      });
      compareContainer.innerHTML = "";
    }
  });
  // --- END ADDED LISTENER ---

  // --- NEW: Show Comparison Report ---
  function showComparisonReport(countryNames) {
    compareContainer.innerHTML = ""; // Clear previous comparison

    // Limit to 2 countries for this layout
    const countriesToCompare = countryNames.slice(0, 2);

    countriesToCompare.forEach((countryName, index) => {
      const countryData = allCountriesData.get(countryName);
      if (!countryData) {
        // Create a column with an error
        const col = document.createElement("div");
        col.className = "compare-column";
        col.innerHTML = `<h3>Country not found: "${countryName}"</h3>`;
        compareContainer.appendChild(col);
        return; // Skip to next country
      }

      // --- Get KPI data for this country ---

      // Use the year-aware normalization helper for sheet discovery
      const normLatest = normalizeDateForSheet(countryData.latestDate);
      const normBench = normalizeDateForSheet(countryData.benchmarkDate);

      const latestSheetName = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes(normLatest.toLowerCase())
      );

      const benchmarkSheetName = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes(normBench.toLowerCase())
      );

      if (!latestSheetName) {
        const col = document.createElement("div");
        col.className = "compare-column";
        col.innerHTML = `<h3>Error</h3><p>Missing data for ${countryData.name}</p>`;
        compareContainer.appendChild(col);
        return;
      }

      const latestKPIs = getKPIsForCountry(latestSheetName, countryData.name);
      const benchmarkKPIs = benchmarkSheetName
        ? getKPIsForCountry(benchmarkSheetName, countryData.name)
        : null;
      // --- End KPI data fetch ---

      // Create a new column element
      const col = document.createElement("div");
      col.className = "compare-column";
      compareContainer.appendChild(col);

      // Generate the report into this column
      updateReportContent(
        col,
        countryData,
        latestKPIs,
        benchmarkKPIs,
        false, // Default to "vs Install"
        `compare${index + 1}`
      );
    });

    // Show the modal
    comparisonModal.style.display = "flex";
    comparisonModal.classList.remove("modal-hidden");
  }
  function displayAAPULists(low, mod, good) {
    [
      [low, lowList],
      [mod, moderateList],
      [good, goodList],
    ].forEach(([data, element]) => {
      element.innerHTML = "";
      if (data.length === 0) {
        element.innerHTML = '<p class="placeholder">No countries found.</p>';
      } else {
        data.forEach((country) => {
          const item = document.createElement("div");

          // 1. Determine Color (Green if New BM, otherwise default Red)
          const isGreen = country.isNewBenchmark;
          item.className = `country-item ${isGreen ? "above" : ""}`;

          // 2. Determine Text (Show "New BM!" if green, or BM value if red)
          let bmText = "";
          if (isGreen) {
            bmText = `🚀 New BM!`;
          } else {
            bmText = `(BM: ${country.benchmarkAAPU.toFixed(2)})`;
          }

          item.innerHTML = `<strong>${
            country.name
          }</strong> <span>${country.latestAAPU.toFixed(
            2
          )} <small>${bmText}</small></span>`;
          item.addEventListener("click", () =>
            isCompareMode
              ? handleCompareSelection(country, item)
              : showWTHReport(country)
          );
          element.appendChild(item);
        });
      }
    });
  }

  // --- GLOBAL TRENDS CORE LOGIC ---

  // 1. The Sheet Crawler: Builds history for all KPIs from every date sheet
  function crawlHistory(wb) {
    const historyMap = {};
    const dateSheets = wb.SheetNames.filter((name) => {
      const parts = name.split("|");
      const datePart = (parts.length > 1 ? parts.pop() : name).trim();
      // Since datePart now contains the year (e.g. "4 Jan 26"), we parse it directly
      return !isNaN(Date.parse(datePart)) && !name.toUpperCase().includes("BM");
    }).sort((a, b) => {
      const dA = a.split("|").pop().trim();
      const dB = b.split("|").pop().trim();
      return Date.parse(dA) - Date.parse(dB);
    });

    const sortedDates = dateSheets.map((s) => s.split("|").pop().trim());

    dateSheets.forEach((sheetName, dateIdx) => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
        header: 1,
      });
      if (!data.length) return;
      const countries = data[0];
      const kpiRows = data.slice(1);

      for (let col = 1; col < countries.length; col++) {
        const country = countries[col];
        if (!country) continue;
        const countryKey = country.toLowerCase();
        if (!historyMap[countryKey]) historyMap[countryKey] = {};

        kpiRows.forEach((row) => {
          const kpi = row[0];
          if (!kpi) return;
          if (!historyMap[countryKey][kpi])
            historyMap[countryKey][kpi] = new Array(sortedDates.length).fill(
              null
            );
          historyMap[countryKey][kpi][dateIdx] =
            typeof row[col] === "number" ? row[col] : null;
        });
      }
    });
    return { dates: sortedDates, data: historyMap };
  }

  // Modern Dropdown Search and Pill Logic (Updated with Groups)
  function populateTrendCountryList() {
    if (!globalCountryList || !activeGameKey) return;
    globalCountryList.innerHTML = "";
    const searchTerm = globalTrendCountrySearch.value.toLowerCase();
    const game = DASHBOARD_STATE[activeGameKey];

    // 1. ADD BATCH GROUPS (Only if not searching)
    if (!searchTerm) {
      const categories = [
        { label: "Group: Low AAPU", key: "low", color: "#fff8f8" },
        { label: "Group: Moderate AAPU", key: "moderate", color: "#fcfbf1" },
        { label: "Group: Good AAPU", key: "good", color: "#f1f8ee" },
      ];

      categories.forEach((cat) => {
        const count = game[cat.key].length;
        if (count === 0) return;

        const groupDiv = document.createElement("div");
        groupDiv.className = `dropdown-item group-item group-${cat.key}`;
        groupDiv.innerHTML = `<span>${cat.label}</span> <span class="count-badge">${count}</span>`;
        groupDiv.onclick = (e) => {
          e.stopPropagation();
          toggleTrendGroup(cat.key);
          globalTrendCountryList.style.display = "none";
        };
        globalCountryList.appendChild(groupDiv);
      });
    }

    // 2. RENDER COUNTRIES
    const sortedNames = Array.from(allCountriesData.keys()).sort();
    let visibleCount = 0;

    sortedNames.forEach((nameKey) => {
      const countryData = allCountriesData.get(nameKey);
      if (!countryData.name.toLowerCase().includes(searchTerm)) return;

      const item = document.createElement("div");
      item.className = `dropdown-item ${
        selectedTrendCountries.has(nameKey) ? "selected" : ""
      }`;
      item.textContent = countryData.name;
      item.onclick = () => {
        toggleTrendCountry(nameKey);
        globalTrendCountryList.style.display = "none";
        globalTrendCountrySearch.value = "";
      };
      globalCountryList.appendChild(item);
      visibleCount++;
    });

    globalTrendCountryList.style.display =
      (visibleCount > 0 || !searchTerm) &&
      document.activeElement === globalTrendCountrySearch
        ? "block"
        : "none";
    renderActivePills();
  }

  // Batch Select Helper
  function toggleTrendGroup(groupKey) {
    const game = DASHBOARD_STATE[activeGameKey];
    const groupCountries = game[groupKey].map((c) => c.name.toLowerCase());

    // Check if ALL countries in this group are already selected
    const allSelected = groupCountries.every((name) =>
      selectedTrendCountries.has(name)
    );

    if (allSelected) {
      // If all selected, deselect them all
      groupCountries.forEach((name) => selectedTrendCountries.delete(name));
    } else {
      // Otherwise, add any missing ones
      groupCountries.forEach((name) => selectedTrendCountries.add(name));
    }

    updateGlobalTrendChart();
    renderActivePills();
  }

  function toggleTrendCountry(nameKey) {
    selectedTrendCountries.has(nameKey)
      ? selectedTrendCountries.delete(nameKey)
      : selectedTrendCountries.add(nameKey);
    updateGlobalTrendChart();
    renderActivePills();
  }

  function renderActivePills() {
    const container = document.getElementById("activeTrendCountries");
    container.innerHTML = "";
    selectedTrendCountries.forEach((nameKey) => {
      const countryData = allCountriesData.get(nameKey);
      const pill = document.createElement("div");
      pill.className = "active-pill";
      pill.innerHTML = `${countryData.name} <span class="remove-btn">×</span>`;
      pill.querySelector(".remove-btn").onclick = () =>
        toggleTrendCountry(nameKey);
      container.appendChild(pill);
    });
  }

  // Handle Focus/Blur for Modern Dropdown
  globalTrendCountrySearch.addEventListener("focus", () => {
    globalTrendCountryList.style.display = "block";
    populateTrendCountryList();
  });

  // 3. Update the Global Chart
  function updateGlobalTrendChart() {
    const canvas = document.getElementById("globalTrendChartCanvas");
    if (!canvas || !activeGameKey) return;

    const gameData = DASHBOARD_STATE[activeGameKey];
    if (!gameData || !gameData.history) return;

    if (chartInstances.has("globalTrend")) {
      chartInstances.get("globalTrend").destroy();
      chartInstances.delete("globalTrend");
    }

    const kpi = activeKPIValue;
    const range = globalDateRangeValue; // Use the new unified state variable

    let dates = [...gameData.history.dates];
    let startIdx = 0;
    let endIdx = dates.length;

    if (range === "7") {
      startIdx = Math.max(0, dates.length - 7);
    } else if (range === "30") {
      startIdx = Math.max(0, dates.length - 30);
    } else if (range === "custom") {
      const sVal = document.getElementById("customDateStart").value;
      const eVal = document.getElementById("customDateEnd").value;
      if (sVal && eVal) {
        const startTs = new Date(sVal).setHours(0, 0, 0, 0);
        const endTs = new Date(eVal).setHours(23, 59, 59, 999);
        // Parse date directly from sheet name format (e.g., "4 Jan 26")
        const timestamps = dates.map((d) => Date.parse(d));

        startIdx = timestamps.findIndex((ts) => ts >= startTs);
        if (startIdx === -1) startIdx = 0;

        const foundEndIdx = timestamps.findIndex((ts) => ts > endTs);
        endIdx = foundEndIdx === -1 ? dates.length : foundEndIdx;
      }
    }

    const sliceDates = dates.slice(startIdx, endIdx);
    const datasets = Array.from(selectedTrendCountries).map((countryKey, i) => {
      const rawValues =
        gameData.history.data[countryKey] &&
        gameData.history.data[countryKey][kpi]
          ? gameData.history.data[countryKey][kpi]
          : [];
      const sliceValues = rawValues.slice(startIdx, endIdx);
      const colors = [
        "#3b82f6",
        "#ef4444",
        "#10b981",
        "#f59e0b",
        "#8b5cf6",
        "#ec4899",
        "#06b6d4",
      ];
      const color = colors[i % colors.length];
      return {
        label: allCountriesData.get(countryKey).name,
        data: sliceValues.map((v) =>
          kpi.includes("%") && v !== null ? v * 100 : v
        ),
        borderColor: color,
        backgroundColor: color, // Set to same as border for solid colored balls
        pointBackgroundColor: color, // Ensures points on the line are solid
        pointRadius: 4, // Makes the balls slightly larger and clearer
        pointHoverRadius: 6,
        tension: 0.2,
        spanGaps: true,
      };
    });

    const isPercentage = kpi.includes("%");
    const ctx = canvas.getContext("2d");

    chartInstances.set(
      "globalTrend",
      new Chart(ctx, {
        type: "line",
        data: { labels: sliceDates, datasets },
        options: {
          layout: {
            padding: { left: 15, right: 15 }, // Gives labels space to grow bold without nudging the axis
          },
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          // 1. TRACKER: Record exactly which date is being touched
          onHover: (event, elements, chart) => {
            chart.hoveredIdx = elements.length > 0 ? elements[0].index : -1;
            chart.update("none"); // Instant update
          },
          scales: {
            x: {
              offset: false, // Line starts exactly at the Y-axis (No more gap)
              afterFit: (scale) => {
                scale.height = 75;
              },
              ticks: {
                // Removed align: 'inner' to prevent horizontal nudging
                minRotation: 45,
                maxRotation: 45,
                padding: 10, // Adds breathing room from the axis line
                color: (ctx) =>
                  ctx.chart.hoveredIdx === ctx.index ? "#000" : "#222", // High contrast base color
                font: (ctx) => {
                  const active = ctx.chart.hoveredIdx === ctx.index;
                  return {
                    family:
                      "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif",
                    weight: active ? "700" : "400",
                    size: 12,
                    style: "normal",
                  };
                },
              },
            },
            y: {
              beginAtZero: true,
              ticks: {
                callback: (v) =>
                  isPercentage ? v.toFixed(2) + "%" : v.toFixed(2),
              },
            },
          },
          plugins: {
            legend: {
              position: "bottom",
              labels: {
                usePointStyle: true,
                pointStyle: "circle",
                padding: 20,
                font: { size: 12, weight: "600" },
              },
            },
            tooltip: {
              backgroundColor: "rgba(17, 24, 39, 0.9)",
              padding: 12,
              titleFont: { size: 14, weight: "bold" },
              bodyFont: { size: 13 },
              usePointStyle: true,
              callbacks: {
                label: (context) => {
                  const countryKey = Array.from(selectedTrendCountries)[
                    context.datasetIndex
                  ];
                  const val = context.parsed.y;
                  const formattedVal = isPercentage
                    ? val.toFixed(2) + "%"
                    : val.toFixed(2);
                  const history = gameData.history.data[countryKey];
                  const dateIdx = startIdx + context.dataIndex;
                  const dailyInstalls =
                    history && history["User Installed"]
                      ? history["User Installed"][dateIdx]
                      : "N/A";

                  return `${context.dataset.label}: ${formattedVal} (Installs: ${dailyInstalls})`;
                },
              },
            },
          },
        },
      })
    );
    // --- CONSOLIDATED RESET: Fixes sticky bold dates on mouse exit ---
    canvas.onmouseout = () => {
      const chart = chartInstances.get("globalTrend");
      if (chart) {
        chart.hoveredIdx = -1; // Resets the correct variable used in tick font logic
        chart.update("none");
      }
    };
  }

  addFileRow();
  // --- Global Trends Event Listeners ---
  if (globalCountrySearch) {
    globalCountrySearch.addEventListener("input", populateTrendCountryList);
  }
  // --- Smart Date Picker Logic ---
  let globalDateRangeValue = "30"; // Tracks active range state
  const dateBtn = document.getElementById("dateRangeDisplay");
  const dateDropdown = document.getElementById("smartDateDropdown");
  const presetBtns = document.querySelectorAll(".preset-btn");
  const customToggle = document.getElementById("customRangeToggle");
  const calendarPanel = document.getElementById("customCalendarPanel");

  if (dateBtn) {
    dateBtn.onclick = (e) => {
      e.stopPropagation();
      const isHidden = dateDropdown.style.display === "none";
      dateDropdown.style.display = isHidden ? "block" : "none";

      if (isHidden) {
        // Sync Light-Theme active classes when opening
        presetBtns.forEach((p) =>
          p.classList.toggle(
            "active",
            p.getAttribute("data-range") === globalDateRangeValue
          )
        );
        if (customToggle)
          customToggle.classList.toggle(
            "active",
            globalDateRangeValue === "custom"
          );
      }
    };
  }

  presetBtns.forEach((btn) => {
    btn.onclick = (e) => {
      e.stopPropagation(); // CRITICAL: Keeps unified popup open when selecting
      const range = btn.getAttribute("data-range");
      if (!range) return;

      document.getElementById("dateRangeText").textContent = btn.textContent;
      globalDateRangeValue = range;
      if (calendarPanel) calendarPanel.style.display = "none";
      if (dateDropdown) dateDropdown.style.display = "none";
      updateGlobalTrendChart();
    };
  });

  if (customToggle) {
    customToggle.onclick = (e) => {
      e.stopPropagation(); // Stops accidental closing
      const isVisible = calendarPanel.style.display === "block";
      calendarPanel.style.display = isVisible ? "none" : "block";
    };
  }

  const applyBtn = document.getElementById("applyCustomDate");
  if (applyBtn) {
    applyBtn.onclick = (e) => {
      e.stopPropagation();
      const s = document.getElementById("customDateStart").value;
      const eVal = document.getElementById("customDateEnd").value;
      if (s && eVal) {
        document.getElementById("dateRangeText").textContent = "Custom Range";
        globalDateRangeValue = "custom";
        dateDropdown.style.display = "none";
        updateGlobalTrendChart();
      } else {
        alert("Please select both dates.");
      }
    };
  }
  const clearTrendBtn = document.getElementById("clearTrendFilters");
  if (clearTrendBtn) {
    clearTrendBtn.onclick = () => {
      selectedTrendCountries.clear();
      if (globalTrendCountrySearch) globalTrendCountrySearch.value = "";
      updateGlobalTrendChart();
      renderActivePills();
    };
  }

  // Ensure clicking inside the calendar doesn't close the picker
  if (calendarPanel) {
    calendarPanel.onclick = (e) => e.stopPropagation();
  }

  // --- GLOBAL MANAGER: Close all dropdowns on outside click ---
  document.addEventListener("click", (e) => {
    const dateWrapper = document.getElementById("unifiedDatePickerWrapper");
    const kpiWrapper = document.getElementById("kpiSelectorWrapper");
    const countryWrapper = document.getElementById("dropdownWrapper");

    // Close Date Window
    if (dateWrapper && !dateWrapper.contains(e.target)) {
      if (dateDropdown) dateDropdown.style.display = "none";
    }
    // Close KPI Selector
    if (kpiWrapper && !kpiWrapper.contains(e.target)) {
      if (kpiDropdown) kpiDropdown.style.display = "none";
    }
    // Close Country Search
    if (countryWrapper && !countryWrapper.contains(e.target)) {
      if (globalTrendCountryList) globalTrendCountryList.style.display = "none";
    }
  });
  // --- Custom KPI Selection Logic (Table View) ---
  const kpiBtn = document.getElementById("kpiDisplayBtn");
  const kpiDropdown = document.getElementById("kpiSmartDropdown");
  const kpiOptions = document.querySelectorAll(".kpi-opt");
  let activeKPIValue = "Avg Ad per user"; // Default initialization

  if (kpiBtn) {
    kpiBtn.onclick = (e) => {
      e.stopPropagation();
      kpiDropdown.style.display =
        kpiDropdown.style.display === "none" ? "block" : "none";
      // Sync UI state
      kpiOptions.forEach((opt) =>
        opt.classList.toggle(
          "active",
          opt.getAttribute("data-value") === activeKPIValue
        )
      );
    };
  }

  kpiOptions.forEach((opt) => {
    opt.onclick = (e) => {
      e.stopPropagation();
      activeKPIValue = opt.getAttribute("data-value");
      document.getElementById("activeKPIText").textContent = opt.textContent;
      kpiDropdown.style.display = "none";
      updateGlobalTrendChart();
    };
  });

  // Modify updateGlobalTrendChart to use this variable
  // Find "const kpi = globalKPISelect.value;" and change to "const kpi = activeKPIValue;" inside updateGlobalTrendChart
  // --- FIX: Close Upload Modal on Background Click ---
  uploadModal.addEventListener("click", (event) => {
    // If the user clicks the dark background (the modal itself), close it
    // But if they click the white box (uploadModal.children[0]), do nothing
    if (event.target === uploadModal) {
      uploadModal.style.display = "none";
      uploadModal.classList.add("modal-hidden");
    }
  });
});
