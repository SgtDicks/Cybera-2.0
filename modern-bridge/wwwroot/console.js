(() => {
  const storageKey = "cybera-bridge-api-key";
  const state = { key: sessionStorage.getItem(storageKey) || "", relayEnabled: false };
  const $ = (id) => document.getElementById(id);
  const dialog = $("key-dialog");
  const keyInput = $("api-key");
  const commandButtons = [...document.querySelectorAll("[data-action]")];

  function setServerState(kind, message) {
    const stateNode = $("server-state");
    stateNode.innerHTML = `<span class="dot ${kind}"></span><span>${message}</span>`;
  }

  function setNotice(title, message) {
    $("integration-notice").innerHTML = `<strong>${title}</strong><span>${message}</span>`;
  }

  async function request(path, options = {}) {
    if (!state.key) throw new Error("Enter an API key to use the protected console.");
    const headers = new Headers(options.headers || {});
    headers.set("X-Api-Key", state.key);
    const response = await fetch(path, { ...options, headers, cache: "no-store" });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || `Request failed (${response.status})`);
    }
    return response.json();
  }

  function hasWorkstations(snapshot) {
    return snapshot && Array.isArray(snapshot.workstations);
  }

  function renderWorkstations(snapshot) {
    const target = $("workstation-list");
    const tag = $("snapshot-tag");
    if (!hasWorkstations(snapshot)) {
      tag.textContent = "No workstation snapshot";
      target.className = "empty-state";
      target.innerHTML = "<strong>Workstation data is not integrated.</strong><p>The source returned no <code>workstations</code> array. This is intentionally different from “zero workstations.”</p>";
      $("workstation-count").textContent = "Unavailable";
      $("workstation-detail").textContent = "No trusted workstation snapshot";
      return;
    }
    tag.textContent = "Snapshot connected";
    target.className = "workstation-list";
    $("workstation-count").textContent = String(snapshot.workstations.length);
    $("workstation-detail").textContent = "Read-only exported snapshot";
    if (snapshot.workstations.length === 0) {
      target.innerHTML = "<div class=\"empty-state\"><strong>Snapshot reports zero workstations.</strong><p>This value comes from the configured status exporter.</p></div>";
      return;
    }
    target.replaceChildren(...snapshot.workstations.map((station) => {
      const card = document.createElement("article"); card.className = "workstation";
      const id = String(station.id ?? station.name ?? "Unidentified workstation");
      const stateText = String(station.state ?? station.session?.state ?? "State unavailable");
      const session = station.session?.id ? `Session ${station.session.id}` : "Session details unavailable";
      card.innerHTML = "";
      const title = document.createElement("h3"); title.textContent = id;
      const detail = document.createElement("p"); detail.textContent = station.name && station.id ? String(station.name) : session;
      const sessionLine = document.createElement("small"); sessionLine.textContent = session;
      const status = document.createElement("p"); status.className = "state"; status.textContent = stateText;
      card.append(title, detail, sessionLine, status);
      return card;
    }));
  }

  function updateCommandAvailability() {
    const hasTarget = /^[A-Za-z0-9._-]{1,64}$/.test($("workstation-id").value);
    commandButtons.forEach((button) => { button.disabled = !(state.key && state.relayEnabled && hasTarget); });
    $("relay-state").textContent = state.relayEnabled ? "Enabled" : "Disabled";
    $("relay-detail").textContent = state.relayEnabled ? "Loopback relay configured" : "No command is sent";
    $("command-note").textContent = state.relayEnabled
      ? (hasTarget ? "Commands are allowlisted and sent only to the local relay." : "Enter a valid workstation ID to enable supported commands.")
      : "The command relay is disabled, so no command can reach Cybera.";
  }

  async function loadFeatures() {
    const target = $("feature-list");
    try {
      const features = await request("/api/v1/features");
      target.className = "workstation-list";
      target.replaceChildren(...features.map((feature) => {
        const card = document.createElement("article"); card.className = "workstation";
        card.innerHTML = "";
        const title = document.createElement("h3"); title.textContent = feature.area;
        const mode = document.createElement("p"); mode.textContent = feature.mode;
        const status = document.createElement("p"); status.className = "state"; status.textContent = feature.status;
        card.append(title, mode, status); return card;
      }));
    } catch {
      target.className = "empty-state";
      target.innerHTML = "<strong>Feature status unavailable.</strong><p>Authenticate the console to inspect the migration matrix.</p>";
    }
  }

  async function refresh() {
    $("command-result").textContent = "";
    if (!state.key) { setServerState("neutral", "API key required"); setNotice("Legacy data not connected.", "Enter an operator API key to inspect protected status."); updateCommandAvailability(); return; }
    try {
      const status = await request("/api/v1/status");
      const running = status.legacyServer?.processRunning === true;
      state.relayEnabled = status.integration?.commandRelayEnabled === true;
      setServerState(running ? "good" : "bad", running ? "Legacy process running" : "Legacy process not found");
      $("legacy-process").textContent = running ? "Running" : "Not running";
      $("legacy-process-detail").textContent = `Local process: ${status.legacyServer?.processName || "unknown"}`;
      renderWorkstations(status.legacyServer?.snapshot);
      setNotice(running ? "Protected status loaded." : "Legacy process was not found.", status.integration?.directLegacyTcp === false ? "The bridge did not open the legacy workstation TCP protocol." : "Review legacy integration configuration.");
      $("last-refresh").textContent = `Updated ${new Date(status.utc).toLocaleTimeString()}`;
      updateCommandAvailability();
      await loadFeatures();
    } catch (error) {
      state.relayEnabled = false; updateCommandAvailability(); setServerState("bad", "Status unavailable");
      setNotice("Could not load protected status.", error.message);
    }
  }

  $("key-button").addEventListener("click", () => { keyInput.value = state.key; dialog.showModal(); keyInput.focus(); });
  $("key-form").addEventListener("submit", (event) => { event.preventDefault(); state.key = keyInput.value.trim(); if (!state.key) return; sessionStorage.setItem(storageKey, state.key); dialog.close(); $("key-button").textContent = "Change API key"; refresh(); });
  $("forget-key").addEventListener("click", () => { state.key = ""; state.relayEnabled = false; sessionStorage.removeItem(storageKey); keyInput.value = ""; dialog.close(); $("key-button").textContent = "Connect API key"; refresh(); });
  $("refresh-button").addEventListener("click", refresh);
  $("workstation-id").addEventListener("input", updateCommandAvailability);
  commandButtons.forEach((button) => button.addEventListener("click", async () => {
    const action = button.dataset.action; const workstation = $("workstation-id").value.trim();
    if (!confirm(`${action === "end" ? "End" : "Pause"} session on ${workstation}?`)) return;
    $("command-result").textContent = "Submitting command…";
    try {
      const command = await request(`/api/v1/sessions/${encodeURIComponent(workstation)}/commands`, { method: "POST", headers: { "Content-Type": "application/json", "Idempotency-Key": crypto.randomUUID() + crypto.randomUUID().slice(0, 16) }, body: JSON.stringify({ action }) });
      $("command-result").textContent = `Command accepted: ${command.id}`;
    } catch (error) { $("command-result").textContent = `Command not sent: ${error.message}`; }
  }));
  $("key-button").textContent = state.key ? "Change API key" : "Connect API key";
  refresh();
})();
