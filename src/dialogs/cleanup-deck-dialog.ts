type ResultType = "success" | "info" | "warning" | "error";

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

  if (runButton) {
    runButton.disabled = isBusy;
    runButton.textContent = isBusy ? "Working..." : "Create Clean Copy";
  }

  if (closeButton) {
    closeButton.disabled = isBusy;
  }
}

function getOptions() {
  const removeComments = (document.getElementById("remove-comments") as HTMLInputElement | null)?.checked ?? false;
  const removeNotes = (document.getElementById("remove-notes") as HTMLInputElement | null)?.checked ?? false;
  return { removeComments, removeNotes };
}

function runCleanup() {
  const options = getOptions();
  if (!options.removeComments && !options.removeNotes) {
    setStatus("warning", "Heads up", "Choose at least one cleanup target before creating a cleaned copy.");
    return;
  }

  setBusy(true);
  setStatus("info", "Working", "Creating a cleaned copy of the current presentation...");
  Office.context.ui.messageParent(JSON.stringify({ type: "run", options }));
}

Office.onReady(() => {
  const runButton = document.getElementById("run");
  const closeButton = document.getElementById("close");

  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg: { message: string }) => {
    try {
      const payload = JSON.parse(arg.message) as {
        type?: string;
        result?: { type: ResultType; message: string };
      };

      if (payload.type !== "result" || !payload.result) {
        return;
      }

      const titleMap: Record<ResultType, string> = {
        success: "Done",
        info: "Info",
        warning: "Heads up",
        error: "Error",
      };

      setStatus(payload.result.type, titleMap[payload.result.type], payload.result.message);
    } catch (error) {
      console.error(error);
      setStatus("error", "Error", "Could not parse the response from the host page.");
    } finally {
      setBusy(false);
    }
  });

  runButton?.addEventListener("click", runCleanup);
  closeButton?.addEventListener("click", () => {
    Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
  });
});
