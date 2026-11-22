let eventQueue = Promise.resolve();
let worksheetListenerRegistered = false;

function setStatus(message) {
  const element = document.getElementById("status");
  if (element) element.textContent = message;
}

function isRelevantAddress(address) {
  const normalized = String(address || "")
    .toUpperCase()
    .replace(/\$/g, "");

  const match = normalized.match(/(?:^|!)([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!match) return false;

  const firstColumn = match[1];
  const firstRow = Number(match[2]);
  const secondColumn = match[3] || firstColumn;
  const secondRow = match[4] ? Number(match[4]) : firstRow;

  const touchesStatusCells =
    firstColumn === "H" &&
    secondColumn === "H" &&
    firstRow <= 17 &&
    secondRow >= 3;

  const touchesResetCell =
    (firstColumn === "F" && firstRow === 18) ||
    (secondColumn === "F" && secondRow === 18) ||
    (firstColumn === "F" && secondColumn === "H" &&
      firstRow <= 18 && secondRow >= 18);

  return touchesStatusCells || touchesResetCell;
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Please open this add-in in Excel.");
    return;
  }

  try {
    // Supported in Excel clients with SharedRuntime 1.1.
    // On clients without this feature, the task pane can still run
    // after the user opens the add-in manually.
    if (Office.addin && typeof Office.addin.setStartupBehavior === "function") {
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

    // X is currently unused in this workbook and stores per-month processing flags.
    ws.getRange("X2").values = [["Processed this month"]];
    ws.getRange("X:X").format.columnHidden = true;

    ws.onChanged.add(handleWorksheetChanged);
    await context.sync();
  });

  worksheetListenerRegistered = true;
  setStatus("Die-Budget Automation is active. Listening for H3:H17 and F18.");
}

function handleWorksheetChanged(event) {
  if (!isRelevantAddress(event.address)) return;

  // Process events sequentially so rapid checkbox changes aren't handled concurrently.
  eventQueue = eventQueue
    .then(() => processBudgetChanges())
    .catch((error) => {
      setStatus("Automation error: " + error.message);
      console.error(error);
    });
}

async function processBudgetChanges() {
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem("Die-Budget");

    // F18 is the anchor cell of the merged F18:H18 control.
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

    // Monthly reset: clear the flags and all 15 checkboxes, but retain balances.
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

      const currentBalance =
        typeof balances[i][0] === "number" ? balances[i][0] : 0;

      // Add G[row] to the existing C[row]; never replace the balance with G[row].
      balances[i][0] = currentBalance + amount;
      processed[i][0] = true;

      itemsProcessed++;
      totalAdded += amount;
    }

    if (itemsProcessed > 0) {
      balancesRange.values = balances;
      processedRange.values = processed;
      await context.sync();

      setStatus(
        `Processed ${itemsProcessed} item(s). Total added: ${totalAdded}.`
      );
    } else {
      setStatus("No new checked items to process.");
    }
  });
}
