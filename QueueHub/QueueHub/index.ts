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
  private _selectedQueue: QueueInfo | null = null;
  private _agents: AgentInfo[] = [];

  private _pollTimer: number | null = null;

  // DOM refs
  private _elSearch!: HTMLInputElement;
  private _elSubtitle!: HTMLDivElement;
  private _elList!: HTMLDivElement;
  private _elBack!: HTMLButtonElement;
  private _elHeader!: HTMLDivElement;
  private _elSummary!: HTMLDivElement;

  private _view: "queues" | "agents" = "queues";

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
      <button class="qh-back" data-ref="back" style="display:none">
        <svg width="14" height="14" viewBox="0 0 20 20" fill="currentColor"><path d="M12.35 3.15a.5.5 0 0 1 0 .7L6.21 10l6.14 6.15a.5.5 0 0 1-.7.7l-6.5-6.5a.5.5 0 0 1 0-.7l6.5-6.5a.5.5 0 0 1 .7 0Z"/></svg>
        <span>All Queues</span>
      </button>
      <div class="qh-header" data-ref="header" style="display:none"></div>
      <div class="qh-summary" data-ref="summary" style="display:none"></div>
      <div class="qh-list" data-ref="list">
        <div class="qh-loading"><span class="qh-loading-dot"></span><span class="qh-loading-dot" style="animation-delay:.2s"></span><span class="qh-loading-dot" style="animation-delay:.4s"></span></div>
      </div>`;

    this._elSearch = this._ref("search") as HTMLInputElement;
    this._elSubtitle = this._ref("subtitle") as HTMLDivElement;
    this._elList = this._ref("list") as HTMLDivElement;
    this._elBack = this._ref("back") as HTMLButtonElement;
    this._elHeader = this._ref("header") as HTMLDivElement;
    this._elSummary = this._ref("summary") as HTMLDivElement;

    this._elSearch.addEventListener("input", () => this._onSearch());
    this._elBack.addEventListener("click", () => this._showQueues());
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
      this._renderQueues();
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

  private _getOrgUrl(): string {
    const xrm = (window as unknown as Record<string, unknown>)["Xrm"] as
      { Utility?: { getGlobalContext?: () => { getClientUrl?: () => string } } } | undefined;
    const url = xrm?.Utility?.getGlobalContext?.()?.getClientUrl?.();
    if (url) return url.replace(/\/+$/, "");
    return "";
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
      });
    }

    if (userIds.length > 0) {
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

    return agents;
  }

  /* ── Rendering: Queue list ── */

  private _renderQueues(filter?: string): void {
    this._view = "queues";
    this._elBack.style.display = "none";
    this._elHeader.style.display = "none";
    this._elSummary.style.display = "none";
    this._elSubtitle.style.display = "";
    this._elSearch.placeholder = "Search queues\u2026";

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
      html += `<div class="qh-item" data-qid="${esc(q.id)}">
        <div class="qh-item-icon">${esc(getInitials(q.name))}</div>
        <div class="qh-item-body">
          <div class="qh-item-name">${esc(q.name)}</div>
        </div>
        <svg class="qh-item-arrow" width="14" height="14" viewBox="0 0 20 20" fill="currentColor"><path d="M7.65 3.15a.5.5 0 0 0 0 .7L13.79 10l-6.14 6.15a.5.5 0 0 0 .7.7l6.5-6.5a.5.5 0 0 0 0-.7l-6.5-6.5a.5.5 0 0 0-.7 0Z"/></svg>
      </div>`;
    }
    this._elList.innerHTML = html;

    this._elList.querySelectorAll(".qh-item").forEach(el => {
      el.addEventListener("click", () => {
        const qid = (el as HTMLElement).dataset.qid!;
        const queue = this._queues.find(q => q.id === qid);
        if (queue) this._openQueue(queue);
      });
    });
  }

  /* ── Rendering: Agent list ── */

  private async _openQueue(queue: QueueInfo): Promise<void> {
    this._selectedQueue = queue;
    this._view = "agents";
    this._elSearch.value = "";
    this._elSearch.placeholder = "Search agents\u2026";
    this._elBack.style.display = "";
    this._elSubtitle.style.display = "none";

    this._elHeader.style.display = "";
    this._elHeader.innerHTML = `
      <div>
        <div class="qh-header-text">${esc(queue.name)}</div>
        <div class="qh-header-count">Loading agents\u2026</div>
      </div>`;
    this._elSummary.style.display = "none";
    this._elSummary.innerHTML = "";
    this._elList.innerHTML = `<div class="qh-loading"><span class="qh-loading-dot"></span><span class="qh-loading-dot" style="animation-delay:.2s"></span><span class="qh-loading-dot" style="animation-delay:.4s"></span></div>`;

    if (this._pollTimer !== null) { clearInterval(this._pollTimer); this._pollTimer = null; }

    try {
      this._agents = await this._loadAgentsInQueue(queue.id);
      this._renderAgents();
      this._pollTimer = window.setInterval(() => this._pollAgents(), POLL_MS);
    } catch (e: unknown) {
      this._elList.innerHTML = `<div class="qh-empty">${esc(e instanceof Error ? e.message : String(e))}</div>`;
    }
  }

  private _renderAgents(filter?: string): void {
    let agents = this._agents;
    if (filter) {
      const lf = filter.toLowerCase();
      agents = agents.filter(a => a.name.toLowerCase().indexOf(lf) > -1);
    }

    const countEl = this._elHeader.querySelector(".qh-header-count");
    if (countEl) countEl.textContent = `${this._agents.length} agent${this._agents.length !== 1 ? "s" : ""}`;

    const totals: Record<string, number> = {};
    for (const a of this._agents) {
      totals[a.presenceName] = (totals[a.presenceName] || 0) + 1;
    }
    const sorted = Object.keys(totals).sort((a, b) => totals[b] - totals[a]);
    let sumHtml = "";
    for (const n of sorted) {
      sumHtml += `<div class="qh-chip"><div class="qh-chip-dot" style="background:${statusColor(n)}"></div><span class="qh-chip-count">${totals[n]}</span><span>${esc(n)}</span></div>`;
    }
    this._elSummary.innerHTML = sumHtml;
    this._elSummary.style.display = sorted.length ? "" : "none";

    if (!agents.length) {
      this._elList.innerHTML = `<div class="qh-empty">${filter ? "No agents match your search" : "No agents in this queue"}</div>`;
      return;
    }

    const statusOrder = (a: AgentInfo): number => {
      const l = a.presenceName.toLowerCase();
      if (l.indexOf("available") > -1) return 0;
      if (l.indexOf("busy") > -1) return 1;
      if (l.indexOf("do not disturb") > -1) return 2;
      if (l.indexOf("away") > -1 || l.indexOf("appear away") > -1) return 3;
      if (l.indexOf("offline") > -1 || l.indexOf("inactive") > -1) return 5;
      return 4;
    };
    agents = [...agents].sort((a, b) => {
      const diff = statusOrder(a) - statusOrder(b);
      return diff !== 0 ? diff : a.name.localeCompare(b.name);
    });

    let html = "";
    for (const a of agents) {
      const col = statusColor(a.presenceName);
      const sinceStr = a.since ? fmtDuration(Date.now() - new Date(a.since).getTime()) : "";
      const isMe = a.id === this._userId;
      const initials = esc(getInitials(a.name));
      const imgUrl = `${this._getOrgUrl()}/api/data/v9.2/systemusers(${a.id})/entityimage/$value`;
      html += `<div class="qh-agent">
        <div class="qh-agent-avatar" style="background:${isMe ? "#e0ecff" : "#f0f0f0"};color:${isMe ? "#0078d4" : "#666"}">
          <img class="qh-agent-photo" src="${imgUrl}" alt="" onload="this.parentElement.style.background='transparent'" onerror="this.style.display='none';this.nextElementSibling.style.display=''" />
          <span class="qh-agent-initials" style="display:none">${initials}</span>
          <div class="qh-agent-dot" style="background:${col}"></div>
        </div>
        <div class="qh-agent-body">
          <div class="qh-agent-name">${esc(a.name)}</div>
          <div class="qh-agent-status"><span style="color:${col}">${esc(a.presenceName)}</span>${sinceStr ? ` \u00b7 ${esc(sinceStr)}` : ""}</div>
        </div>
        ${isMe ? '<span class="qh-agent-you">You</span>' : ""}
      </div>`;
    }
    this._elList.innerHTML = html;
  }

  /* ── Polling ── */

  private async _pollAgents(): Promise<void> {
    if (this._view !== "agents" || !this._selectedQueue) return;
    try {
      this._agents = await this._loadAgentsInQueue(this._selectedQueue.id);
      this._renderAgents(this._elSearch.value.trim() || undefined);
    } catch {
      // silently retry next poll
    }
  }

  /* ── Search ── */

  private _onSearch(): void {
    const val = this._elSearch.value.trim();
    if (this._view === "queues") {
      this._renderQueues(val || undefined);
    } else {
      this._renderAgents(val || undefined);
    }
  }

  /* ── Navigation ── */

  private _showQueues(): void {
    if (this._pollTimer !== null) { clearInterval(this._pollTimer); this._pollTimer = null; }
    this._selectedQueue = null;
    this._agents = [];
    this._elSearch.value = "";
    this._renderQueues();
  }
}
