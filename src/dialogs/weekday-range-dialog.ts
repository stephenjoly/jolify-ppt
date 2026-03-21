type ResultType = "success" | "info" | "warning" | "error";

type DialogResultPayload = {
  type: "result";
  result: {
    type: ResultType;
    message: string;
    output: string;
  };
};

function setStatus(type: ResultType, title: string, message: string) {
  const el = document.getElementById("status");
  if (!el) {
    return;
  }

  el.className = `status ${type}`;
  el.innerHTML = "";

  const strong = document.createElement("strong");
  strong.textContent = title;
  el.appendChild(strong);

  const span = document.createElement("span");
  span.textContent = message;
  el.appendChild(span);
}

function setBusy(isBusy: boolean) {
  const runButton = document.getElementById("run") as HTMLButtonElement | null;
  const closeButton = document.getElementById("close") as HTMLButtonElement | null;
  const copyButton = document.getElementById("copy") as HTMLButtonElement | null;

  if (runButton) {
    runButton.disabled = isBusy;
    runButton.textContent = isBusy ? "Working..." : "Generate";
  }

  if (closeButton) {
    closeButton.disabled = isBusy;
  }

  if (copyButton) {
    copyButton.disabled = isBusy;
  }
}

function setResults(value: string) {
  const textarea = document.getElementById("results") as HTMLTextAreaElement | null;
  if (textarea) {
    textarea.value = value;
  }
}

function getInputValue(id: string) {
  return (document.getElementById(id) as HTMLInputElement | HTMLSelectElement | null)?.value ?? "";
}

function runGenerate() {
  const startDate = getInputValue("start-date");
  const endDate = getInputValue("end-date");
  const weekday = Number(getInputValue("weekday"));

  setBusy(true);
  setStatus("info", "Working", "Generating the weekday list...");

  Office.context.ui.messageParent(
    JSON.stringify({
      type: "run",
      params: {
        startDate,
        endDate,
        weekday,
      },
    }),
  );
}

async function copyResults() {
  const textarea = document.getElementById("results") as HTMLTextAreaElement | null;
  if (!textarea?.value) {
    setStatus("info", "Info", "Nothing to copy yet.");
    return;
  }

  try {
    await navigator.clipboard.writeText(textarea.value);
    setStatus("success", "Copied", "Copied the generated dates to the clipboard.");
  } catch (error) {
    console.error(error);
    textarea.focus();
    textarea.select();
    setStatus("warning", "Heads up", "Could not access the clipboard automatically. The results are selected so you can copy them manually.");
  }
}

Office.onReady(() => {
  const runButton = document.getElementById("run");
  const closeButton = document.getElementById("close");
  const copyButton = document.getElementById("copy");

  const today = new Date().toISOString().slice(0, 10);
  const end = document.getElementById("end-date") as HTMLInputElement | null;
  const start = document.getElementById("start-date") as HTMLInputElement | null;
  if (start) {
    start.value = today;
  }
  if (end) {
    end.value = today;
  }

  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg: { message: string }) => {
    try {
      const payload = JSON.parse(arg.message) as DialogResultPayload;
      if (payload.type !== "result") {
        return;
      }

      const titleMap: Record<ResultType, string> = {
        success: "Done",
        info: "Info",
        warning: "Heads up",
        error: "Error",
      };

      setStatus(payload.result.type, titleMap[payload.result.type], payload.result.message);
      setResults(payload.result.output);
    } catch (error) {
      console.error(error);
      setStatus("error", "Error", "Could not parse the response from the host page.");
    } finally {
      setBusy(false);
    }
  });

  runButton?.addEventListener("click", () => {
    runGenerate();
  });

  copyButton?.addEventListener("click", () => {
    void copyResults();
  });

  closeButton?.addEventListener("click", () => {
    Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
  });
});
