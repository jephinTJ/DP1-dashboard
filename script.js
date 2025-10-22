// --- Wait for the page to be fully loaded ---
document.addEventListener("DOMContentLoaded", () => {
  // Get references to all THREE list elements
  const fileUploader = document.getElementById("fileUploader");
  const criticalList = document.getElementById("criticalList");
  const belowList = document.getElementById("belowList");
  const aboveList = document.getElementById("aboveList");
  const modal = document.getElementById("reportModal");
  const modalContent = document.getElementById("reportContent");
  const closeModalButton = document.getElementById("closeModalButton");
  const mainTitle = document.getElementById("mainTitle");

  let workbook = null; // This will store the loaded Excel file
  let myTrendChart = null; // This will hold the chart object

  // --- Listen for a file upload ---
  fileUploader.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    // Reset all three lists
    criticalList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    belowList.innerHTML = '<p class="placeholder">Analyzing file...</p>';
    aboveList.innerHTML = '<p class="placeholder">Analyzing file...</p>';

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });

      // Start the analysis
      processBenchmarkSheet(file.name);
    };
    reader.readAsArrayBuffer(file);
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
    let gameName = "DIQ2"; // Default to DIQ2
    if (fileName && fileName.toUpperCase().includes("DIQ-3")) {
      gameName = "DIQ-3";
    } // Add more 'else if' conditions here if you have other game names like DIQ-4 etc.

    // --- Update Title using Game Name and Date ---
    mainTitle.textContent = `${gameName} AAPU Performance Dashboard of ${latestDate}`;
    document.title = `AAPU Report | ${gameName} | ${latestDate}`; // Update browser tab too

    const criticalCountries = [];
    const belowCountries = [];
    const aboveCountries = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0 || !row[0]) continue;

      const country = row[0];
      const latestValue = row[latestDateCol];

      // *** ADD THIS BLOCK START ***
      // Get installs from the latest sheet for this country
      const latestSheetName = workbook.SheetNames.find((name) =>
        name.includes(latestDate)
      );
      let userInstalls = 0;
      if (latestSheetName) {
        const latestKPIs = getKPIsForCountry(latestSheetName, country); // Use helper
        if (latestKPIs && latestKPIs["User Installed"] !== undefined) {
          // Make sure installs is a number, handle potential 'NA' or errors
          userInstalls = Number(latestKPIs["User Installed"]) || 0;
        }
      }

      // --- FILTER CONDITION ---
      if (userInstalls <= 10) {
        continue; // Skip this country if installs are 10 or less
      }
      // *** ADD THIS BLOCK END ***

      let benchmarkValue = -1.0;
      let benchmarkDate = "N_A";

      for (let j = 1; j < latestDateCol; j++) {
        const currentValue = row[j];
        if (typeof currentValue === "number" && currentValue > benchmarkValue) {
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
    }
    displayCountryLists(criticalCountries, belowCountries, aboveCountries);
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

      item.addEventListener("click", () => showWTHReport(country));
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

  // --- Function to show WTH report ---
  function showWTHReport(country) {
    if (myTrendChart) {
      myTrendChart.destroy();
    }

    const latestSheetName = workbook.SheetNames.find((name) =>
      name.includes(country.latestDate)
    );
    const benchmarkSheetName = workbook.SheetNames.find((name) =>
      name.includes(country.benchmarkDate)
    );

    if (!latestSheetName) {
      modalContent.innerHTML = `<h3>Error</h3><p>Could not find data sheet for latest date: "${country.latestDate}"</p>`;
      modal.style.display = "flex";
      modal.classList.remove("modal-hidden");
      return;
    }
    // Handle case where benchmark date might be very old and not have a sheet?
    if (!benchmarkSheetName && country.benchmarkDate !== "N_A") {
      modalContent.innerHTML = `<h3>Error</h3><p>Could not find data sheet for benchmark date: "${country.benchmarkDate}"</p>`;
      modal.style.display = "flex";
      modal.classList.remove("modal-hidden");
      return;
    }

    const latestKPIs = getKPIsForCountry(latestSheetName, country.name);
    // Only get benchmark KPIs if the sheet exists
    const benchmarkKPIs = benchmarkSheetName
      ? getKPIsForCountry(benchmarkSheetName, country.name)
      : null;

    const reportHTML = generateReportHTML(country, latestKPIs, benchmarkKPIs);
    modalContent.innerHTML = reportHTML;

    const trendChartCanvas = document.getElementById("trendChart");
    if (!trendChartCanvas) {
      console.error("Chart canvas not found!");
      return; // Don't try to draw if canvas isn't there
    }

    const ctx = trendChartCanvas.getContext("2d");
    myTrendChart = new Chart(ctx, {
      type: "line",
      data: {
        labels: country.trendDates,
        datasets: [
          {
            label: `Last ${country.trendDates.length} Days AAPU`, // Dynamic label
            data: country.trendValuesFormatted, // Use formatted for chart
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
          legend: {
            display: false,
          },
          tooltip: {
            mode: "index",
            intersect: false,
          },
        },
        scales: {
          y: {
            beginAtZero: false,
          },
        },
      },
    });

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
  function generateReportHTML(country, latest, bench) {
    // Bench might be null if benchmark sheet wasn't found (e.g., very old benchmark)
    if (!latest)
      return `<h3>Error</h3><p>Missing KPI data for ${country.name} on ${country.latestDate}.</p>`;

    const isDown = country.latestAAPU < country.benchmarkAAPU;
    const statusClass = isDown ? "down" : "up";

    // --- 1. KPI 'What Happened' Section ---
    const kpiMap = [
      { key: "% of users at 20", name: "Retention at Lvl 20", type: "ret" },
      { key: "% of users at 50", name: "Retention at Lvl 50", type: "ret" },
      { key: "% of users at 70", name: "Retention at Lvl 70", type: "ret" },
      { key: "% of users at 100", name: "Retention at Lvl 100", type: "ret" },
      { key: "% of users at 150", name: "Retention at Lvl 150", type: "ret" },
      { key: "% of users at 200", name: "Retention at Lvl 200", type: "ret" },
      { key: "% of users at Ads 10", name: "Users at Ads 10", type: "ad" },
      { key: "% of users at Ads 20", name: "Users at Ads 20", type: "ad" },
      { key: "% of users at Ads 40", name: "Users at Ads 40", type: "ad" },
      { key: "% of users at Ads 70", name: "Users at Ads 70", type: "ad" },
      { key: "% of users at Ads 100", name: "Users at Ads 100", type: "ad" },
    ];

    const retKPIs = kpiMap.filter((k) => k.type === "ret");
    const adKPIs = kpiMap.filter((k) => k.type === "ad");

    // --- Helper function to generate list items ---
    const generateListItems = (kpiList) => {
      // *** NEW: Check if it's a new benchmark day ***
      if (country.isNewBenchmark) {
        // If new benchmark, just show the current values without comparison
        return kpiList
          .map((k) => {
            const l = latest[k.key] || 0;
            // Only show current value
            return `<li>${k.name}: <strong>${(l * 100).toFixed(
              2
            )}%</strong> <small>(New BM)</small></li>`;
          })
          .join("");
      } else {
        // Original logic for comparison if NOT a new benchmark
        return kpiList
          .map((k) => {
            const l = latest[k.key] || 0;
            const b = bench ? bench[k.key] || 0 : 0;
            const change_pp = bench ? (l - b) * 100 : 0;

            let cssClass = "";
            let kpiWarningSymbol = "";

            if (bench && change_pp < -1.0) {
              cssClass = "down";
              if (change_pp < -4.0) {
                kpiWarningSymbol = " ❗";
              }
            } else if (bench && change_pp > 1.0) {
              cssClass = "up";
            }

            const bmText = bench ? `(${(b * 100).toFixed(2)}% BM)` : "(BM N/A)";

            return `<li>${k.name}: <strong class="${cssClass}">${(
              l * 100
            ).toFixed(
              2
            )}%${kpiWarningSymbol}</strong> <small>${bmText}</small></li>`;
          })
          .join("");
      }
    }; // --- End of generateListItems ---

    const retListHTML = generateListItems(retKPIs);
    const adListHTML = generateListItems(adKPIs);

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

    // --- 4. Assemble the Report ---
    return `
            <div class="wth-report">
                <h2>Country Performance: ${country.name}</h2>
                <p class="summary">
    <strong>Current AAPU: <span class="${statusClass}">${country.latestAAPU.toFixed(
      2
    )}</span></strong> 
    ${
      country.isNewBenchmark
        ? `(🚀 New Benchmark!)`
        : `(Benchmark: ${country.benchmarkAAPU.toFixed(2)} on ${
            country.benchmarkDate
          })`
    }
</p>
                <h3>${whatHappenedTitle}</h3>
                <div class="kpi-columns">
                    <div class="kpi-column">
                        <h4>Retention KPIs</h4>
                        <ul>${retListHTML}</ul>
                    </div>
                    <div class="kpi-column">
                        <h4>Ad KPIs</h4>
                        <ul>${adListHTML}</ul>
                    </div>
                </div>
                <h3>Trend Analysis (last ${country.totalDays} days)</h3> 
                ${trendSummary}
                <canvas id="trendChart" height="150"></canvas>
                <h3>${analysisTitle}</h3>
                ${analysisAdvice}
            </div>
        `;
  }

  // --- Close modal ---
  function closeModal() {
    modal.style.display = "none";
    modal.classList.add("modal-hidden");
    if (myTrendChart) {
      myTrendChart.destroy();
      myTrendChart = null;
    }
  }

  closeModalButton.addEventListener("click", closeModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) {
      closeModal();
    }
  });
});
