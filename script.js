// --- Wait for the page to be fully loaded ---
document.addEventListener("DOMContentLoaded", () => {
  // Get references to all THREE list elements
  const fileUploader = document.getElementById("fileUploader");
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
  // --- End New Elements ---

  let workbook = null; // This will store the loaded Excel file
  let chartInstances = new Map(); // Replaces myTrendChart

  // --- Data for Single Modal ---
  let currentCountryData = null;
  let currentLatestKPIs = null;
  let currentBenchmarkKPIs = null;
  // --- End Data ---

  // --- State Variables ---
  let isCompareMode = false;
  let allCountriesData = new Map(); // Stores all country data for search
  let selectedCountriesForCompare = []; // Array to hold selected country names

  // --- Listen for a file upload ---
  fileUploader.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    // Clear compare mode data
    allCountriesData.clear();
    compareToggle.checked = false;
    isCompareMode = false;
    countrySearchInput.placeholder = "Search for a country...";
    countrySearchInput.value = "";

    // Reset all three lists
    criticalList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    belowList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    aboveList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    lowList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    moderateList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    goodList.innerHTML = '<p class="placeholder">Analyzing file...</p>';

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });

      // Start the analysis
      processBenchmarkSheet(file.name);
    };
    reader.readAsArrayBuffer(file);
  });
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

  // --- Function to process the 'BM' sheet ---
  function processBenchmarkSheet(fileName) {
    const bmSheetName = workbook.SheetNames.find(
      (name) => name.toUpperCase() === "BM"
    );
    if (!bmSheetName) {
      criticalList.innerHTML =
        '<p class="placeholder" style="color: red;">Error: "BM" sheet not found.</p>';
      belowList.innerHTML =
        '<p class="placeholder" style="color: red;">Error: "BM" sheet not found.</p>';
      aboveList.innerHTML =
        '<p class="placeholder" style="color: red;">Error: "BM" sheet not found.</p>';
      return;
    }

    const ws = workbook.Sheets[bmSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const headers = data[0];
    const latestDate = headers[headers.length - 1];
    const latestDateCol = headers.length - 1;
    const totalDays = latestDateCol; // *** Total number of date columns ***

    // --- Determine Game Name from Filename ---
    let gameName = "DIQ-2"; // Default to DIQ2
    if (fileName && fileName.toUpperCase().includes("DIQ-3")) {
      gameName = "DIQ-3";
    } // Add more 'else if' conditions here if you have other game names like DIQ-4 etc.

    // --- Update Title using Game Name and Date ---
    mainTitle.innerHTML = `${gameName} AAPU Performance Dashboard of ${latestDate} <span style="font-size: 0.6em; color: #737373; vertical-align: middle;">(Last ${totalDays} Days)</span>`;
    document.title = `AAPU Report | ${gameName} | ${latestDate}`; // Update browser tab too

    const criticalCountries = [];
    const belowCountries = [];
    const aboveCountries = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0 || !row[0]) continue;

      let country = row[0]; // <-- Declare and assign初步 here

      try {
        // country = row[0]; // Assign inside try block - No longer needed here
        const latestValue = row[latestDateCol];

        // *** ADD THIS BLOCK START ***
        // Get installs from the latest sheet for this country
        const latestSheetName = workbook.SheetNames.find((name) =>
          name.includes(String(latestDate))
        );
        console.log(`Date: ${latestDate}, Found Sheet: ${latestSheetName}`);
        // --- START: ADDED ERROR HANDLING ---
        if (!latestSheetName) {
          console.error(
            `Error: Could not find data sheet including date "${latestDate}" for country "${country}". Skipping this country.`
          );
          continue; // Skip to the next country row
        }
        // --- END: ADDED ERROR HANDLING ---
        let userInstalls = 0;
        if (latestSheetName) {
          const latestKPIs = getKPIsForCountry(latestSheetName, country); // Use helper
          if (latestKPIs && latestKPIs["User Installed"] !== undefined) {
            // Make sure installs is a number, handle potential 'NA' or errors
            userInstalls = Number(latestKPIs["User Installed"]) || 0;
          }
        }

        console.log(`Country: ${country}, Installs: ${userInstalls}`);
        // --- FILTER CONDITION ---
        if (userInstalls <= 10) {
          continue; // Skip this country if installs are 10 or less
        }
        // *** ADD THIS BLOCK END ***

        let benchmarkValue = -1.0;
        let benchmarkDate = "N_A";

        for (let j = 1; j < latestDateCol; j++) {
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

        const numHeaders = headers.length;
        const startCol = Math.max(1, numHeaders - 7); // Use 7 for 7 days
        const trendDates = headers.slice(startCol, numHeaders);
        // Store raw numbers for trend calculation
        const trendValuesRaw = row
          .slice(startCol, numHeaders)
          .map((v) => (typeof v === "number" ? v : null));
        // For chart display
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
            trendValuesRaw: trendValuesRaw, // Use raw for calculation
            trendValuesFormatted: trendValuesFormatted, // Use formatted for chart
            totalDays: totalDays, // *** Pass total days ***
            isNewBenchmark: isNewBenchmark,
          };
          // --- Store in Map for search ---
          allCountriesData.set(country.toLowerCase(), countryData);

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
        // --- START: ADDED CATCH BLOCK ---
      } catch (error) {
        console.error(
          `Error processing country: ${country || "Unknown"} (Row ${i + 1})`,
          error
        );
        // Optionally add this country to a list of errors to display later
      }
      // --- END: ADDED CATCH BLOCK ---
    } // This is the closing brace for the main 'for' loop
    displayCountryLists(criticalCountries, belowCountries, aboveCountries);
    // --- NEW AAPU LOGIC ---
    const low = [],
      moderate = [],
      good = [];
    allCountriesData.forEach((c) => {
      if (c.latestAAPU < 7) low.push(c);
      else if (c.latestAAPU <= 8) moderate.push(c);
      else good.push(c);
    });
    displayAAPULists(low, moderate, good);
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

    // --- 1. Find sheets ---
    const latestSheetName = workbook.SheetNames.find((name) =>
      name.includes(country.latestDate)
    );
    const benchmarkSheetName = workbook.SheetNames.find((name) =>
      name.includes(country.benchmarkDate)
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

    const latestKPIs = getKPIsForCountry(latestSheetName, country.name);
    const benchmarkKPIs = benchmarkSheetName
      ? getKPIsForCountry(benchmarkSheetName, country.name)
      : null;

    // --- 2. Store data globally for toggle ---
    currentCountryData = country;
    currentLatestKPIs = latestKPIs;
    currentBenchmarkKPIs = benchmarkKPIs;

    // --- 3. Set Defaults and Initial Display ---
    reportCountryNameSpan.textContent = country.name; // Set country name in header

    // *** NEW: Update the static summary in the header ***
    const summaryP = document.querySelector(
      ".modal-content .report-header p.summary"
    );
    if (summaryP) {
      const isDown =
        !country.isNewBenchmark && country.latestAAPU < country.benchmarkAAPU;
      const statusClass = isDown ? "down" : "up";
      const benchmarkFullText = country.isNewBenchmark
        ? `(🚀 New Benchmark!)`
        : `(Benchmark: ${country.benchmarkAAPU.toFixed(2)} on ${
            country.benchmarkDate
          })`;
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
  function getKPIsForCountry(sheetName, countryName) {
    const ws = workbook.Sheets[sheetName];
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

    // --- 2. *** CORRECTED Trend Analysis Logic *** ---
    let consecutiveDays = 0;
    let trendSummary = "";
    let trendSummaryText = "";
    const trendValues = country.trendValuesRaw;
    const n = trendValues.length;
    const trendDaysCount = country.totalDays;
    const benchmarkAAPU = country.benchmarkAAPU;

    // *** ADD THIS IF CHECK ***
    if (country.isNewBenchmark) {
      trendSummaryText = `🚀 New Benchmark!`;
      trendSummary = `<p>${trendSummaryText}</p>`;
    }
    // *** END ADDED IF CHECK ***
    // Add an 'else' around the original logic
    else if (n > 0) {
      if (isDown) {
        // ... (rest of the original 'isDown' logic remains unchanged inside this else if) ...
        for (let i = n - 1; i >= 0; i--) {
          const value = trendValues[i];
          if (value === null || value < benchmarkAAPU) {
            consecutiveDays++;
          } else {
            break;
          }
        }
        if (
          consecutiveDays === 1 &&
          n > 1 &&
          trendValues[n - 2] !== null &&
          trendValues[n - 2] >= benchmarkAAPU
        ) {
          trendSummaryText = `Dropped below BM today`;
        } else if (consecutiveDays > 0) {
          trendSummaryText = `Below BM for last ${consecutiveDays} consecutive day${
            consecutiveDays > 1 ? "s" : ""
          }`;
        } else {
          trendSummaryText = `Status unclear (check recent data)`;
        }
        trendSummary = `<p>${trendSummaryText}.</p>`;
      } else {
        // AAPU is UP or Equal (but not a new benchmark today)
        // ... (rest of the original 'else' logic remains unchanged inside this block) ...
        for (let i = n - 1; i >= 0; i--) {
          const value = trendValues[i];
          if (value !== null && value >= benchmarkAAPU) {
            consecutiveDays++;
          } else {
            break;
          }
        }
        if (
          consecutiveDays === 1 &&
          n > 1 &&
          (trendValues[n - 2] === null || trendValues[n - 2] < benchmarkAAPU)
        ) {
          trendSummaryText = `Rose to/above BM today`;
        } else if (consecutiveDays > 0) {
          trendSummaryText = `At/Above BM for last ${consecutiveDays} consecutive day${
            consecutiveDays > 1 ? "s" : ""
          }`;
        } else {
          trendSummaryText = `Status unclear (check recent data)`;
        }
        trendSummary = `<p>${trendSummaryText}.</p>`;
      }
    } else {
      // n <= 0
      trendSummaryText = "Not enough data";
      trendSummary = `<p>${trendSummaryText}.</p>`;
    }
    // --- END Trend Analysis Logic ---

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
        : `(Benchmark: ${country.benchmarkAAPU.toFixed(2)} on ${
            country.benchmarkDate
          })`;

      finalDetailsHTML += `
        <div class="report-header">
            <h2>Country Performance: ${country.name}</h2>
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
            <h3>Trend Analysis (last ${reportData.trendDaysCount} days)</h3> 
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
            labels: country.trendDates, // <-- USES LOCAL country VARIABLE
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

      // --- Get KPI data for this country (logic from showWTHReport) ---
      const latestSheetName = workbook.SheetNames.find((name) =>
        name.includes(countryData.latestDate)
      );
      const benchmarkSheetName = workbook.SheetNames.find((name) =>
        name.includes(countryData.benchmarkDate)
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
      // We give it a unique ID 'compare1' or 'compare2'
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
});
