import { IInputs, IOutputs } from "./generated/ManifestTypes";

type FilterPayload = {
  mode: "SR" | "BU";
  field: "vp_salesrepresentative" | "vp_businessunit";
  values: string[];
  capApplied: boolean;
};

export class VITAFilterInjector implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context!: ComponentFramework.Context<IInputs>;
  private hasRun = false;

  public init(ctx: ComponentFramework.Context<IInputs>): void {
    this.context = ctx;
    // small delay to allow grid to mount
    setTimeout(() => { void this.run(); }, 600);
  }

  public updateView(): void { /* no-op */ }
  public getOutputs(): IOutputs { return {}; }
  public destroy(): void { /* no-op */ }

  // ---------------- core ----------------

  private async run(): Promise<void> {
    if (this.hasRun) return;
    this.hasRun = true;

    const debug = await this.readEnv("vp_VITA_FilterInjection_Debug");
    const enabled = await this.readEnv("vp_VITA_EnableFilterInjection");
    if (!enabled) { if (debug) this.log("disabled"); return; }

    const entity = this.resolveEntity();
    if (!entity) { if (debug) this.log("no entity"); return; }

    const payload = await this.fetchFilterSet(entity);
    if (!payload?.values?.length) { if (debug) this.log("empty payload"); return; }

    if (debug) {
      this.log(`entity=${entity} mode=${payload.mode} field=${payload.field} count=${payload.values.length} cap=${payload.capApplied}`);
      this.log(`values=${payload.values.join(",")}`);
    }

    const grid: any = this.findGridControl();
    if (!grid || typeof grid.getFetchXml !== "function" || typeof grid.setFilterXml !== "function") {
      if (debug) this.log("grid control not found");
      return;
    }

    try {
      const base: string = grid.getFetchXml();
      const updated = this.injectInCondition(base, payload.field, payload.values, debug);
      if (updated && updated !== base) {
        grid.setFilterXml(updated);
        if (typeof grid.refresh === "function") grid.refresh();
        if (debug) this.log("filter applied");
      } else {
        if (debug) this.log("no change");
      }
    } catch (e: any) {
      if (debug) this.log(`inject error: ${e?.message || e}`);
    }
  }

  // ---------------- helpers: env / api / entity / grid ----------------

  private async readEnv(schema: string): Promise<boolean> {
    const webApi = (this.context as any).webAPI as ComponentFramework.WebApi;
    const res = await webApi.retrieveMultipleRecords(
      "environmentvariabledefinition",
      `?$select=schemaname,defaultvalue&$filter=schemaname eq '${schema}'&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`
    );
    if (!res.entities?.length) return false;

    const def: any = res.entities[0];
    const val = def.environmentvariabledefinition_environmentvariablevalue?.[0]?.value ?? def.defaultvalue;
    const s = (val || "").toString().trim().toLowerCase();
    return s === "true" || s === "yes" || s === "1";
  }

  private resolveEntity(): string | null {
    try {
      const pg = (window as any).Xrm?.Utility?.getPageContext?.();
      const etn = pg?.input?.entityName;
      if (etn) return etn.toLowerCase();
    } catch { /* ignore */ }

    try {
      const url = new URL(window.location.href);
      const etn = url.searchParams.get("etn");
      if (etn) return etn.toLowerCase();
    } catch { /* ignore */ }

    return null;
  }

  private async fetchFilterSet(entity: string): Promise<FilterPayload | null> {
    const webApi = (this.context as any).webAPI as ComponentFramework.WebApi;
    const anyApi: any = webApi;
    let resp: any;

    if (typeof anyApi.invokeUnboundFunction === "function") {
      resp = await anyApi.invokeUnboundFunction("vita_GetViewFilterSet", { EntityLogicalName: entity });
    } else {
      // fallback for older runtimes
      resp = await webApi.retrieveMultipleRecords(
        `vita_GetViewFilterSet(EntityLogicalName='${encodeURIComponent(entity)}')`,
        ""
      );
    }

    const json = resp?.Response ?? resp?.entities?.[0]?.Response;
    if (!json) return null;

    try { return JSON.parse(json) as FilterPayload; } catch { return null; }
  }

  private findGridControl(): any | null {
    try {
      const pc: any[] = (window as any).Xrm?.App?.getVisibleControls?.() || [];
      const g = pc.find(c => typeof c?.getFetchXml === "function" && typeof c?.setFilterXml === "function");
      if (g) return g;
    } catch { /* ignore */ }

    try {
      const controls: any[] = (window as any).Xrm?.Page?.ui?.controls?.get() || [];
      const g = controls.find(c => typeof c?.getFetchXml === "function" && typeof c?.setFilterXml === "function");
      if (g) return g;
    } catch { /* ignore */ }

    return null;
  }

  // ---------------- XML inject (robust typing) ----------------

  private injectInCondition(fetchXml: string, field: string, values: string[], debug: boolean): string {
    const xmlDoc = new DOMParser().parseFromString(fetchXml, "text/xml");

    const fetchEl = xmlDoc.getElementsByTagName("fetch")[0] as Element | undefined;
    if (!fetchEl) return fetchXml;

    const entityEl = fetchEl.getElementsByTagName("entity")[0] as Element | undefined;
    if (!entityEl) return fetchXml;

    // All direct <filter> children of <entity>
    const directFilters = Array.from(entityEl.getElementsByTagName("filter"))
      .filter(f => f.parentElement === entityEl) as Element[];

    // Find existing root AND filter
    let hostAnd: Element | null =
      directFilters.find(f => (f.getAttribute("type") || "").toLowerCase() === "and") || null;

    // If missing, create it and move current direct filters under it
    if (!hostAnd) {
      hostAnd = xmlDoc.createElement("filter") as unknown as Element;
      hostAnd.setAttribute("type", "and");

      for (const f of directFilters) {
        entityEl.removeChild(f);
        hostAnd.appendChild(f);
      }
      entityEl.appendChild(hostAnd);
    }

    // Safety: still null? then bail
    if (!hostAnd) return fetchXml;

    // Remove any direct conditions on the same attribute (we will replace)
    const directConds = Array
      .from(hostAnd.children)
      .filter(n => n.nodeName.toLowerCase() === "condition")
      .map(n => n as Element)
      .filter(c => (c.getAttribute("attribute") || "").toLowerCase() === field.toLowerCase());

    // If an identical IN already exists, skip (cheap compare)
    const existingIn = directConds.find(c => (c.getAttribute("operator") || "").toLowerCase() === "in");
    if (existingIn) {
      const existingVals = Array.from(existingIn.getElementsByTagName("value"))
        .map(v => (v.textContent || "").toLowerCase());

      const want = [...new Set(values.map(v => v.toLowerCase()))].sort();
      const have = [...new Set(existingVals)].sort();

      if (want.length === have.length && want.every((v, i) => v === have[i])) {
        if (debug) this.log("identical IN found → skip");
        return fetchXml;
      }
    }

    for (const c of directConds) hostAnd.removeChild(c);

    // Comment marker (optional)
    hostAnd.appendChild(xmlDoc.createComment(` VITA-INJECT field=${field} ts=${new Date().toISOString()} `));

    // Build new IN condition
    const cond = xmlDoc.createElement("condition") as unknown as Element;
    cond.setAttribute("attribute", field);
    cond.setAttribute("operator", "in");

    for (const g of values) {
      const v = xmlDoc.createElement("value") as unknown as Element;
      v.textContent = g;
      cond.appendChild(v);
    }

    hostAnd.appendChild(cond);

    // Serialize back to string
    return new XMLSerializer().serializeToString(xmlDoc);
  }

  private log(m: string): void {
    try { console.log(`[VITA-PCF] ${m}`); } catch { /* ignore */ }
  }
}
