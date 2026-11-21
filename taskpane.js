
let eventQueue = Promise.resolve();

function setStatus(message) {
  const element = document.getElementById("status");
  if (element) element.textContent = message;
}

function isRelevantAddress(address) {
  const normalized = address.toUpperCase().replace(/\$/g, "");

  // Monthly reset checkbox (merged F18:H18)
  if (normalized.includes("F18")) return true;

  // Checkboxes H3:H17
  return /(?:^|[!:])H(?:[3-9]|1[0-7])(?=$|:)/.test(normalized);
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Please open this add-in in Excel.");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem("Die-Budget");

      // Reserve column X for duplicate prevention
      ws.getRange("X2").values = [["Processed"]];
      ws.getRange("X:X").format.columnHidden = true;

      // Register automatic change listener
      ws.onChanged.add(handleWorksheetChanged);

      await context.sync();
    });

    setStatus("Die-Budget Automation is active. Listening for checkbox changes.");

  } catch (error) {
    setStatus("Initialization error: " + error.message);
    console.error(error);
  }
});

function handleWorksheetChanged(event) {
  if (!isRelevantAddress(event.address)) return;

  // Queue events to prevent simultaneous processing
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

    const resetCell = ws.getRange("F18");
    const statusRange = ws.getRange("H3:H17");
    const amountRange = ws.getRange("G3:G17");
    const balanceRange = ws.getRange("C3:C17");
    const processedRange = ws.getRange("X3:X17");

    resetCell.load("values");
    statusRange.load("values");
    amountRange.load("values");
    balanceRange.load("values");
    processedRange.load("values");

    await context.sync();

    // 1. MONTHLY RESET
    if (resetCell.values[0][0] === true) {

      statusRange.values =
        Array.from({ length: 15 }, () => [false]);

      processedRange.clear(Excel.ClearApplyTo.contents);

      resetCell.values = [[false]];

      await context.sync();

      setStatus("Monthly reset completed successfully.");
      return;
    }

    // 2. PROCESS CHECKED ITEMS
    const statuses = statusRange.values;
    const amounts = amountRange.values;
    const balances = balanceRange.values;
    const processed = processedRange.values;

    let totalAdded = 0;
    let itemsProcessed = 0;

    for (let i = 0; i < 15; i++) {

      if (
        statuses[i][0] === true &&
        processed[i][0] !== true
      ) {
        const amount = amounts[i][0];
        const currentBalance = balances[i][0];

        if (
          typeof amount !== "number" ||
          !Number.isFinite(amount)
        ) {
          continue;
        }

        const current =
          typeof currentBalance === "number"
            ? currentBalance
            : 0;

        // ADD instead of replacing
        balances[i][0] = current + amount;

        // Mark as already processed
        processed[i][0] = true;

        totalAdded += amount;
        itemsProcessed++;
      }
    }

    if (itemsProcessed > 0) {
      balanceRange.values = balances;
      processedRange.values = processed;

      await context.sync();

      setStatus(
        `${itemsProcessed} item(s) processed. Total added: ${totalAdded}`
      );
    } else {
      setStatus("No new checked items to process.");
    }
  });
}
