let eventQueue = Promise.resolve();
let worksheetListenerRegistered = false;

function setStatus(message) {
  const element = document.getElementById("status");
  if (element) element.textContent = message;
}

function isRelevantAddress(address) {
  const normalized = String(address || "").toUpperCase().replace(/\$/g, "");
  if (/(^|!)F18(?::H18)?$/.test(normalized)) return true;
  return /(^|!)H(?:[3-9]|1[0-7])$/.test(normalized);
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Please open this add-in in Excel for the web.");
    return;
  }

  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }

    await registerWorksheetListener();
  } catch (error) {
    setStatus("Initialization error: " + error.message);
    console.error(error);
  }
});

async function registerWorksheetListener() {
  if (worksheetListenerRegistered) return;

  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem("Die-Budget");
    ws.onChanged.add(handleWorksheetChanged);
    await context.sync();
  });

  worksheetListenerRegistered = true;
  setStatus("Die-Budget Automation is active. Change a checkbox in H3:H17 or F18.");
}

function handleWorksheetChanged(event) {
  if (!isRelevantAddress(event.address)) return;

  eventQueue = eventQueue
    .then(processBudgetChanges)
    .catch((error) => {
      setStatus("Automation error: " + error.message);
      console.error(error);
    });
}

async function processBudgetChanges() {
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem("Die-Budget");
    const resetCell = ws.getRange("F18");
    const statusesRange = ws.getRange("H3:H17");
    const amountsRange = ws.getRange("G3:G17");
    const balancesRange = ws.getRange("C3:C17");
    const processedRange = ws.getRange("X3:X17");

    resetCell.load("values");
    statusesRange.load("values");
    amountsRange.load("values");
    balancesRange.load("values");
    processedRange.load("values");
    await context.sync();

    if (resetCell.values[0][0] === true) {
      statusesRange.values = Array.from({ length: 15 }, () => [false]);
      processedRange.clear(Excel.ClearApplyTo.contents);
      resetCell.values = [[false]];
      await context.sync();
      setStatus("Monthly reset completed. Existing balances were preserved.");
      return;
    }

    const statuses = statusesRange.values;
    const amounts = amountsRange.values;
    const balances = balancesRange.values;
    const processed = processedRange.values;

    let itemsProcessed = 0;
    let totalAdded = 0;

    for (let i = 0; i < 15; i++) {
      if (statuses[i][0] !== true || processed[i][0] === true) continue;

      const amount = amounts[i][0];
      if (typeof amount !== "number" || !Number.isFinite(amount)) continue;

      const current = typeof balances[i][0] === "number" ? balances[i][0] : 0;
      balances[i][0] = current + amount;
      processed[i][0] = true;
      itemsProcessed++;
      totalAdded += amount;
    }

    if (itemsProcessed > 0) {
      balancesRange.values = balances;
      processedRange.values = processed;
      await context.sync();
      setStatus(`Processed ${itemsProcessed} item(s). Total added: ${totalAdded}.`);
    } else {
      setStatus("No new checked items to process.");
    }
  });
}
