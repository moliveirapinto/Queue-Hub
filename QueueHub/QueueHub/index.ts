import { IInputs, IOutputs } from "./generated/ManifestTypes";

const POLL_MS = 10000;

const COLORS: Record<string, string> = {
  available: "#92c353",
  busy: "#c4314b",
  "busy - dnd": "#c4314b",
  "do not disturb": "#c4314b",
  away: "#fcd116",
  "appear away": "#fcd116",
  offline: "#8c8c8c",
  inactive: "#8c8c8c",
  "busy - after conversation work": "#e3008c",
  "after conversation work": "#e3008c",
  "dnd-initiating outbound call": "#c4314b",
  "voice consult dnd": "#c4314b",
  "do not disturb - after conversation work": "#e3008c",
};

function statusColor(name: string): string {
  const l = (name || "").toLowerCase();
  for (const k of Object.keys(COLORS)) {
    if (l.indexOf(k) > -1) return COLORS[k];
  }
  return "#8c8c8c";
}

function fmtDuration(ms: number): string {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  if (h > 0) return `${h}h ${m}m`;
  if (m > 0) return `${m}m ${s % 60}s`;
  return `${s % 60}s`;
}

function getInitials(name: string): string {
  const parts = (name || "?").trim().split(/\s+/);
  if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  return parts[0][0].toUpperCase();
}

function esc(s: string): string {
  const t = document.createElement("span");
  t.textContent = s;
  return t.innerHTML;
}

function statusOrder(a: AgentInfo): number {
  const l = a.presenceName.toLowerCase();
  if (l.indexOf("available") > -1) return 0;
  if (l.indexOf("busy") > -1) return 1;
  if (l.indexOf("do not disturb") > -1) return 2;
  if (l.indexOf("away") > -1 || l.indexOf("appear away") > -1) return 3;
  if (l.indexOf("offline") > -1 || l.indexOf("inactive") > -1) return 5;
  return 4;
}

const LOADING_HTML = `<div class="qh-loading"><span class="qh-loading-dot"></span><span class="qh-loading-dot" style="animation-delay:.2s"></span><span class="qh-loading-dot" style="animation-delay:.4s"></span></div>`;

interface QueueInfo {
  id: string;
  name: string;
}

interface AgentInfo {
  id: string;
  name: string;
  presenceId: string | null;
  presenceName: string;
  since: string | null;
  photo: string | null;
}

interface AgentWithQueues extends AgentInfo {
  queues: QueueInfo[];
}

interface WebApiLike {
  retrieveMultipleRecords: (
    entity: string,
    query: string,
    maxPageSize?: number
  ) => Promise<ComponentFramework.WebApi.RetrieveMultipleResponse>;
}

export class QueueHub implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _container!: HTMLDivElement;
  private _context!: ComponentFramework.Context<IInputs>;

  private _userId: string | null = null;
  private _pmap: Record<string, string> = {};
  private _queues: QueueInfo[] = [];

  // Queues subtab state
  private _selectedQueueIds = new Set<string>();
  private _queueAgents: AgentInfo[] = [];

  // Agents subtab state
  private _allAgents: AgentWithQueues[] = [];
  private _selectedAgentIds = new Set<string>();
  private _expandedAgentIds = new Set<string>();

  private _pollTimer: number | null = null;

  // DOM refs
  private _elSearch!: HTMLInputElement;
  private _elSubtitle!: HTMLDivElement;
  private _elList!: HTMLDivElement;
  private _elSummary!: HTMLDivElement;
  private _elTabQueues!: HTMLButtonElement;
  private _elTabAgents!: HTMLButtonElement;

  private _activeTab: "queues" | "agents" = "queues";
  private _dataLoaded = false;

  constructor() {
    // empty
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    _notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._container = container;
    this._container.classList.add("queue-hub");
    try {
      this._buildUI();
      this._initialize();
    } catch (e: unknown) {
      this._container.textContent = `Init error: ${e instanceof Error ? e.message : String(e)}`;
    }
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    if (this._pollTimer !== null) clearInterval(this._pollTimer);
  }

  /* ── UI ── */

  private _buildUI(): void {
    this._container.innerHTML = `
      <div class="qh-search-wrap">
        <input class="qh-search" data-ref="search" placeholder="Search queues\u2026" autocomplete="off" />
      </div>
      <div class="qh-subtitle" data-ref="subtitle">Monitor your team\u2019s availability across every queue you belong to<br>\u2014 in real time \u2014</div>
      <div class="qh-tabs" data-ref="tabs">
        <button class="qh-tab qh-tab--active" data-ref="tab-queues" data-tab="queues">Queues</button>
        <button class="qh-tab" data-ref="tab-agents" data-tab="agents">Agents</button>
      </div>
      <div class="qh-summary" data-ref="summary" style="display:none"></div>
      <div class="qh-list" data-ref="list">${LOADING_HTML}</div>`;

    this._elSearch = this._ref("search") as HTMLInputElement;
    this._elSubtitle = this._ref("subtitle") as HTMLDivElement;
    this._elList = this._ref("list") as HTMLDivElement;
    this._elSummary = this._ref("summary") as HTMLDivElement;
    this._elTabQueues = this._ref("tab-queues") as HTMLButtonElement;
    this._elTabAgents = this._ref("tab-agents") as HTMLButtonElement;

    this._elSearch.addEventListener("input", () => this._onSearch());
    this._elTabQueues.addEventListener("click", () => this._switchTab("queues"));
    this._elTabAgents.addEventListener("click", () => this._switchTab("agents"));
  }

  private _ref(name: string): HTMLElement {
    return this._container.querySelector(`[data-ref="${name}"]`) as HTMLElement;
  }

  /* ── Init ── */

  private async _initialize(): Promise<void> {
    try {
      this._userId = this._getUserId();
      await this._loadPresenceMap();
      await this._loadQueues();
      this._dataLoaded = true;
      this._renderQueuesTab();
    } catch (e: unknown) {
      this._elList.innerHTML = `<div class="qh-empty">${esc(e instanceof Error ? e.message : String(e))}</div>`;
    }
  }

  /* ── Data helpers ── */

  private _getWebApi(): WebApiLike {
    if (this._context.webAPI) return this._context.webAPI;
    const xrm = (window as unknown as Record<string, unknown>)["Xrm"] as
      { WebApi?: WebApiLike } | undefined;
    if (xrm?.WebApi) return xrm.WebApi;
    throw new Error("WebAPI not available");
  }

  private _getClientUrl(): string {
    const xrm = (window as unknown as Record<string, unknown>)["Xrm"] as
      { Utility?: { getGlobalContext?: () => { getClientUrl?: () => string } } } | undefined;
    const url = xrm?.Utility?.getGlobalContext?.()?.getClientUrl?.();
    if (url) return url;
    return window.location.origin;
  }

  private _getUserId(): string {
    const ctx = this._context as ComponentFramework.Context<IInputs> & { userSettings?: { userId?: string } };
    const uid = ctx.userSettings?.userId;
    if (uid) return uid.replace(/[{}]/g, "").toLowerCase();
    const xrm = (window as unknown as Record<string, unknown>)["Xrm"] as
      { Utility?: { getGlobalContext?: () => { userSettings?: { userId?: string } } } } | undefined;
    const xrmUid = xrm?.Utility?.getGlobalContext?.()?.userSettings?.userId;
    if (xrmUid) return xrmUid.replace(/[{}]/g, "").toLowerCase();
    throw new Error("Cannot determine user ID");
  }

  private async _loadPresenceMap(): Promise<void> {
    const api = this._getWebApi();
    const resp = await api.retrieveMultipleRecords(
      "msdyn_presence",
      "?$select=msdyn_presenceid,msdyn_presencestatustext"
    );
    for (const e of resp.entities) {
      this._pmap[e.msdyn_presenceid as string] = e.msdyn_presencestatustext as string;
    }
  }

  private _presenceName(id: string | null): string {
    if (!id) return "Unknown";
    return this._pmap[id] || "Unknown";
  }

  /* ── Queue loading ── */

  private async _loadQueues(): Promise<void> {
    const api = this._getWebApi();
    const fetchXml = `<fetch>
      <entity name="queue">
        <attribute name="queueid"/>
        <attribute name="name"/>
        <order attribute="name"/>
        <link-entity name="queuemembership" from="queueid" to="queueid" intersect="true">
          <link-entity name="systemuser" from="systemuserid" to="systemuserid">
            <filter><condition attribute="systemuserid" operator="eq" value="${this._userId}"/></filter>
          </link-entity>
        </link-entity>
      </entity>
    </fetch>`;
    const encoded = encodeURIComponent(fetchXml);
    const resp = await api.retrieveMultipleRecords("queue", `?fetchXml=${encoded}`, 5000);
    this._queues = [];
    for (const e of resp.entities) {
      const name = e["name"] as string;
      if (/^<.*>$/.test(name) || /^[0-9a-f]{20,}_\d+$/i.test(name)) continue;
      this._queues.push({ id: e["queueid"] as string, name });
    }
  }

  /* ── Agents in queue ── */

  private async _loadAgentsInQueue(queueId: string): Promise<AgentInfo[]> {
    const api = this._getWebApi();
    const fetchXml = `<fetch>
      <entity name="systemuser">
        <attribute name="systemuserid"/>
        <attribute name="fullname"/>
        <order attribute="fullname"/>
        <link-entity name="queuemembership" from="systemuserid" to="systemuserid" intersect="true">
          <link-entity name="queue" from="queueid" to="queueid">
            <filter><condition attribute="queueid" operator="eq" value="${queueId}"/></filter>
          </link-entity>
        </link-entity>
      </entity>
    </fetch>`;
    const encoded = encodeURIComponent(fetchXml);
    const resp = await api.retrieveMultipleRecords("systemuser", `?fetchXml=${encoded}`, 5000);

    const agents: AgentInfo[] = [];
    const userIds: string[] = [];
    for (const e of resp.entities) {
      const uid = e["systemuserid"] as string;
      userIds.push(uid);
      agents.push({
        id: uid,
        name: (e["fullname"] as string) || "Unknown",
        presenceId: null,
        presenceName: "Unknown",
        since: null,
        photo: null,
      });
    }

    await this._enrichPresence(agents, userIds);
    return agents;
  }

  private async _enrichPresence(agents: AgentInfo[], userIds: string[]): Promise<void> {
    if (userIds.length === 0) return;
    const api = this._getWebApi();
    const statusMap: Record<string, { presenceId: string | null; since: string | null }> = {};
    const batchSize = 10;
    for (let i = 0; i < userIds.length; i += batchSize) {
      const batch = userIds.slice(i, i + batchSize);
      const filter = batch.map(id => `_msdyn_agentid_value eq ${id}`).join(" or ");
      try {
        const sr = await api.retrieveMultipleRecords(
          "msdyn_agentstatus",
          `?$filter=${filter}&$select=_msdyn_agentid_value,_msdyn_currentpresenceid_value,msdyn_presencemodifiedon`
        );
        for (const s of sr.entities) {
          const aid = s["_msdyn_agentid_value"] as string;
          statusMap[aid] = {
            presenceId: (s["_msdyn_currentpresenceid_value"] as string) || null,
            since: (s["msdyn_presencemodifiedon"] as string) || null,
          };
        }
      } catch {
        // Some agents may not have status records
      }
    }
    for (const a of agents) {
      const st = statusMap[a.id];
      if (st) {
        a.presenceId = st.presenceId;
        a.presenceName = this._presenceName(st.presenceId);
        a.since = st.since;
      }
    }
  }

  /* ── Load all agents across all queues (for Agents tab) ── */

  private async _loadAllAgentsWithQueues(): Promise<AgentWithQueues[]> {
    const agentMap: Record<string, AgentWithQueues> = {};
    for (const q of this._queues) {
      const agents = await this._loadAgentsInQueue(q.id);
      for (const a of agents) {
        if (!agentMap[a.id]) {
          agentMap[a.id] = { ...a, queues: [] };
        } else {
          // Update presence if newer
          if (a.presenceId) {
            agentMap[a.id].presenceId = a.presenceId;
            agentMap[a.id].presenceName = a.presenceName;
            agentMap[a.id].since = a.since;
          }
        }
        agentMap[a.id].queues.push(q);
      }
    }
    return Object.values(agentMap).sort((a, b) => a.name.localeCompare(b.name));
  }

  /* ── Load agents for multiple selected queues ── */

  private async _loadAgentsForSelectedQueues(): Promise<AgentInfo[]> {
    const agentMap: Record<string, AgentInfo> = {};
    for (const qid of this._selectedQueueIds) {
      const agents = await this._loadAgentsInQueue(qid);
      for (const a of agents) {
        if (!agentMap[a.id]) {
          agentMap[a.id] = { ...a };
        } else if (a.presenceId) {
          agentMap[a.id].presenceId = a.presenceId;
          agentMap[a.id].presenceName = a.presenceName;
          agentMap[a.id].since = a.since;
        }
      }
    }
    return Object.values(agentMap).sort((a, b) => {
      const diff = statusOrder(a) - statusOrder(b);
      return diff !== 0 ? diff : a.name.localeCompare(b.name);
    });
  }

  /* ── Tab switching ── */

  private _switchTab(tab: "queues" | "agents"): void {
    if (this._activeTab === tab) return;
    if (this._pollTimer !== null) { clearInterval(this._pollTimer); this._pollTimer = null; }
    this._activeTab = tab;
    this._elSearch.value = "";
    this._elSummary.style.display = "none";
    this._elSummary.innerHTML = "";

    this._elTabQueues.classList.toggle("qh-tab--active", tab === "queues");
    this._elTabAgents.classList.toggle("qh-tab--active", tab === "agents");

    if (tab === "queues") {
      this._elSearch.placeholder = "Search queues\u2026";
      this._renderQueuesTab();
    } else {
      this._elSearch.placeholder = "Search agents\u2026";
      this._renderAgentsTab();
    }
  }

  /* ══════════════════════════════════
     QUEUES SUBTAB
     ══════════════════════════════════ */

  private _renderQueuesTab(filter?: string): void {
    let queues = this._queues;
    if (filter) {
      const lf = filter.toLowerCase();
      queues = queues.filter(q => q.name.toLowerCase().indexOf(lf) > -1);
    }

    if (!queues.length) {
      this._elList.innerHTML = `<div class="qh-empty">${filter ? "No queues match your search" : "No queues found"}</div>`;
      return;
    }

    let html = "";
    for (const q of queues) {
      const checked = this._selectedQueueIds.has(q.id);
      html += `<div class="qh-item qh-item--selectable${checked ? " qh-item--selected" : ""}" data-qid="${esc(q.id)}">
        <label class="qh-checkbox-wrap" onclick="event.stopPropagation()">
          <input type="checkbox" class="qh-checkbox" data-qid="${esc(q.id)}" ${checked ? "checked" : ""} />
          <span class="qh-checkbox-custom"></span>
        </label>
        <div class="qh-item-icon">${esc(getInitials(q.name))}</div>
        <div class="qh-item-body">
          <div class="qh-item-name">${esc(q.name)}</div>
        </div>
      </div>`;
    }

    // Agents results section
    if (this._selectedQueueIds.size > 0) {
      html += `<div class="qh-results-divider">
        <span>Agents in selected queues</span>
      </div>`;
      html += `<div data-ref="queue-agents">${LOADING_HTML}</div>`;
    }

    this._elList.innerHTML = html;

    // Bind checkbox events
    this._elList.querySelectorAll(".qh-checkbox").forEach(cb => {
      cb.addEventListener("change", (ev) => {
        const input = ev.target as HTMLInputElement;
        const qid = input.dataset.qid!;
        if (input.checked) {
          this._selectedQueueIds.add(qid);
        } else {
          this._selectedQueueIds.delete(qid);
        }
        this._renderQueuesTab(this._elSearch.value.trim() || undefined);
        if (this._selectedQueueIds.size > 0) {
          this._loadAndRenderQueueAgents();
        }
      });
    });

    // Also allow clicking the row itself to toggle
    this._elList.querySelectorAll(".qh-item--selectable").forEach(el => {
      el.addEventListener("click", () => {
        const qid = (el as HTMLElement).dataset.qid!;
        const cb = el.querySelector(".qh-checkbox") as HTMLInputElement;
        cb.checked = !cb.checked;
        cb.dispatchEvent(new Event("change"));
      });
    });

    // If queues are selected, load agents
    if (this._selectedQueueIds.size > 0) {
      this._loadAndRenderQueueAgents();
    }
  }

  private async _loadAndRenderQueueAgents(): Promise<void> {
    const target = this._container.querySelector('[data-ref="queue-agents"]') as HTMLDivElement;
    if (!target) return;

    try {
      this._queueAgents = await this._loadAgentsForSelectedQueues();
      this._renderQueueAgentsSection(target);
      // Start polling
      if (this._pollTimer !== null) clearInterval(this._pollTimer);
      this._pollTimer = window.setInterval(() => this._pollQueueAgents(), POLL_MS);
    } catch (e: unknown) {
      target.innerHTML = `<div class="qh-empty">${esc(e instanceof Error ? e.message : String(e))}</div>`;
    }
  }

  private _renderQueueAgentsSection(target: HTMLElement): void {
    const agents = this._queueAgents;
    // Summary chips
    const totals: Record<string, number> = {};
    for (const a of agents) {
      totals[a.presenceName] = (totals[a.presenceName] || 0) + 1;
    }
    const sorted = Object.keys(totals).sort((a, b) => totals[b] - totals[a]);
    let sumHtml = `<div class="qh-summary" style="display:${sorted.length ? "flex" : "none"}">`;
    for (const n of sorted) {
      sumHtml += `<div class="qh-chip"><div class="qh-chip-dot" style="background:${statusColor(n)}"></div><span class="qh-chip-count">${totals[n]}</span><span>${esc(n)}</span></div>`;
    }
    sumHtml += `</div>`;

    if (!agents.length) {
      target.innerHTML = `${sumHtml}<div class="qh-empty">No agents in selected queues</div>`;
      return;
    }

    let html = sumHtml;
    html += `<div class="qh-results-count">${agents.length} agent${agents.length !== 1 ? "s" : ""}</div>`;
    for (const a of agents) {
      html += this._agentCardHtml(a);
    }
    target.innerHTML = html;
  }

  private async _pollQueueAgents(): Promise<void> {
    if (this._activeTab !== "queues" || this._selectedQueueIds.size === 0) return;
    try {
      this._queueAgents = await this._loadAgentsForSelectedQueues();
      const target = this._container.querySelector('[data-ref="queue-agents"]') as HTMLDivElement;
      if (target) this._renderQueueAgentsSection(target);
    } catch {
      // silently retry
    }
  }

  /* ══════════════════════════════════
     AGENTS SUBTAB
     ══════════════════════════════════ */

  private async _renderAgentsTab(filter?: string): Promise<void> {
    if (!this._allAgents.length || !this._dataLoaded) {
      this._elList.innerHTML = LOADING_HTML;
      try {
        this._allAgents = await this._loadAllAgentsWithQueues();
      } catch (e: unknown) {
        this._elList.innerHTML = `<div class="qh-empty">${esc(e instanceof Error ? e.message : String(e))}</div>`;
        return;
      }
    }

    let agents = this._allAgents;
    if (filter) {
      const lf = filter.toLowerCase();
      agents = agents.filter(a => a.name.toLowerCase().indexOf(lf) > -1);
    }

    if (!agents.length) {
      this._elList.innerHTML = `<div class="qh-empty">${filter ? "No agents match your search" : "No agents found"}</div>`;
      return;
    }

    // Sort by presence then name
    agents = [...agents].sort((a, b) => {
      const diff = statusOrder(a) - statusOrder(b);
      return diff !== 0 ? diff : a.name.localeCompare(b.name);
    });

    let html = "";
    for (const a of agents) {
      const expanded = this._expandedAgentIds.has(a.id);
      const col = statusColor(a.presenceName);
      const sinceStr = a.since ? fmtDuration(Date.now() - new Date(a.since).getTime()) : "";
      const isMe = a.id === this._userId;
      const initials = esc(getInitials(a.name));
      const imgUrl = `${this._getClientUrl()}/api/data/v9.2/systemusers(${a.id})/entityimage/$value`;
      const photoHtml = `<img class="qh-agent-photo" src="${esc(imgUrl)}" alt="" onload="this.parentElement.style.background='transparent'" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'" /><span class="qh-agent-initials" style="display:none">${initials}</span>`;

      html += `<div class="qh-agent-expandable${expanded ? " qh-agent-expandable--open" : ""}" data-aid="${esc(a.id)}">
        <div class="qh-agent qh-agent--clickable">
          <div class="qh-agent-avatar" style="background:${isMe ? "#e0ecff" : "#f0f0f0"};color:${isMe ? "#0078d4" : "#666"}">
            ${photoHtml}
            <div class="qh-agent-dot" style="background:${col}"></div>
          </div>
          <div class="qh-agent-body">
            <div class="qh-agent-name">${esc(a.name)}</div>
            <div class="qh-agent-status"><span style="color:${col}">${esc(a.presenceName)}</span>${sinceStr ? ` \u00b7 ${esc(sinceStr)}` : ""}</div>
          </div>
          ${isMe ? '<span class="qh-agent-you">You</span>' : ""}
          <span class="qh-agent-queue-count">${a.queues.length} queue${a.queues.length !== 1 ? "s" : ""}</span>
          <svg class="qh-expand-arrow${expanded ? " qh-expand-arrow--open" : ""}" width="14" height="14" viewBox="0 0 20 20" fill="currentColor"><path d="M15.85 7.65a.5.5 0 0 0-.7 0L10 12.79 4.85 7.65a.5.5 0 0 0-.7.7l5.5 5.5a.5.5 0 0 0 .7 0l5.5-5.5a.5.5 0 0 0 0-.7Z"/></svg>
        </div>
        <div class="qh-agent-queues" style="display:${expanded ? "block" : "none"}">
          ${a.queues.map(q => `<div class="qh-agent-queue-item">
            <div class="qh-item-icon qh-item-icon--sm">${esc(getInitials(q.name))}</div>
            <span>${esc(q.name)}</span>
          </div>`).join("")}
        </div>
      </div>`;
    }

    this._elList.innerHTML = html;

    // Bind expand/collapse
    this._elList.querySelectorAll(".qh-agent--clickable").forEach(el => {
      el.addEventListener("click", () => {
        const wrapper = el.closest(".qh-agent-expandable") as HTMLElement;
        const aid = wrapper.dataset.aid!;
        const queuesDiv = wrapper.querySelector(".qh-agent-queues") as HTMLElement;
        const arrow = wrapper.querySelector(".qh-expand-arrow") as HTMLElement;
        if (this._expandedAgentIds.has(aid)) {
          this._expandedAgentIds.delete(aid);
          queuesDiv.style.display = "none";
          wrapper.classList.remove("qh-agent-expandable--open");
          arrow.classList.remove("qh-expand-arrow--open");
        } else {
          this._expandedAgentIds.add(aid);
          queuesDiv.style.display = "block";
          wrapper.classList.add("qh-agent-expandable--open");
          arrow.classList.add("qh-expand-arrow--open");
        }
      });
    });

    // Start polling for agents tab
    if (this._pollTimer !== null) clearInterval(this._pollTimer);
    this._pollTimer = window.setInterval(() => this._pollAgentsTab(), POLL_MS);
  }

  private async _pollAgentsTab(): Promise<void> {
    if (this._activeTab !== "agents") return;
    try {
      this._allAgents = await this._loadAllAgentsWithQueues();
      this._renderAgentsTab(this._elSearch.value.trim() || undefined);
    } catch {
      // silently retry
    }
  }

  /* ── Shared agent card HTML ── */

  private _agentCardHtml(a: AgentInfo): string {
    const col = statusColor(a.presenceName);
    const sinceStr = a.since ? fmtDuration(Date.now() - new Date(a.since).getTime()) : "";
    const isMe = a.id === this._userId;
    const initials = esc(getInitials(a.name));
    const imgUrl = `${this._getClientUrl()}/api/data/v9.2/systemusers(${a.id})/entityimage/$value`;
    const photoHtml = `<img class="qh-agent-photo" src="${esc(imgUrl)}" alt="" onload="this.parentElement.style.background='transparent'" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'" /><span class="qh-agent-initials" style="display:none">${initials}</span>`;
    return `<div class="qh-agent">
      <div class="qh-agent-avatar" style="background:${isMe ? "#e0ecff" : "#f0f0f0"};color:${isMe ? "#0078d4" : "#666"}">
        ${photoHtml}
        <div class="qh-agent-dot" style="background:${col}"></div>
      </div>
      <div class="qh-agent-body">
        <div class="qh-agent-name">${esc(a.name)}</div>
        <div class="qh-agent-status"><span style="color:${col}">${esc(a.presenceName)}</span>${sinceStr ? ` \u00b7 ${esc(sinceStr)}` : ""}</div>
      </div>
      ${isMe ? '<span class="qh-agent-you">You</span>' : ""}
    </div>`;
  }

  /* ── Search ── */

  private _onSearch(): void {
    const val = this._elSearch.value.trim();
    if (this._activeTab === "queues") {
      this._renderQueuesTab(val || undefined);
    } else {
      this._renderAgentsTab(val || undefined);
    }
  }
}
