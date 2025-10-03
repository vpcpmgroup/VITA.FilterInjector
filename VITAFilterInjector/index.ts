import { IInputs, IOutputs } from "./generated/ManifestTypes";

type FilterPayload = { mode: "SR"|"BU"; field: "vp_salesrepresentative"|"vp_businessunit"; values: string[]; capApplied: boolean; };

export class VITAFilterInjector implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context!: ComponentFramework.Context<IInputs>;
  private hasRun = false;

  public init(ctx: ComponentFramework.Context<IInputs>) { this.context = ctx; setTimeout(()=>{void this.run();}, 600); }
  public updateView() { /* no-op */ }
  public getOutputs(): IOutputs { return {}; }
  public destroy() { /* no-op */ }

  private async run(): Promise<void> {
    if (this.hasRun) return; this.hasRun = true;

    const debug = await this.readEnv("vp_VITA_FilterInjection_Debug");
    const enabled = await this.readEnv("vp_VITA_EnableFilterInjection");
    if (!enabled) { if (debug) this.log("disabled"); return; }

    const entity = this.resolveEntity(); if (!entity) { if (debug) this.log("no entity"); return; }

    const payload = await this.fetchFilterSet(entity);
    if (!payload?.values?.length) { if (debug) this.log("empty payload"); return; }
    if (debug) { this.log(`entity=${entity} mode=${payload.mode} field=${payload.field} count=${payload.values.length} cap=${payload.capApplied}`);
                 this.log(`values=${payload.values.join(",")}`); }

    const grid: any = this.findGridControl();
    if (!grid || typeof grid.getFetchXml!=="function" || typeof grid.setFilterXml!=="function") { if (debug) this.log("grid control not found"); return; }

    try {
      const base = grid.getFetchXml();
      const updated = this.injectInCondition(base, payload.field, payload.values, debug);
      if (updated && updated !== base) {
        grid.setFilterXml(updated);
        if (typeof grid.refresh==="function") grid.refresh();
        if (debug) this.log("filter applied");
      } else { if (debug) this.log("no change"); }
    } catch(e:any){ if (debug) this.log(`inject error: ${e?.message||e}`); }
  }

  private async readEnv(schema: string): Promise<boolean> {
    const webApi = (this.context as any).webAPI as ComponentFramework.WebApi;
    const res = await webApi.retrieveMultipleRecords("environmentvariabledefinition",
      `?$select=schemaname,defaultvalue&$filter=schemaname eq '${schema}'&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`);
    if (!res.entities?.length) return false;
    const def:any = res.entities[0];
    const val = def.environmentvariabledefinition_environmentvariablevalue?.[0]?.value ?? def.defaultvalue;
    const s = (val||"").toString().trim().toLowerCase();
    return s==="true"||s==="yes"||s==="1";
  }

  private resolveEntity(): string|null {
    try { const pg = (window as any).Xrm?.Utility?.getPageContext?.(); const etn = pg?.input?.entityName; if (etn) return etn.toLowerCase(); } catch {}
    try { const url = new URL(window.location.href); const etn = url.searchParams.get("etn"); if (etn) return etn.toLowerCase(); } catch {}
    return null;
  }

  private async fetchFilterSet(entity: string): Promise<FilterPayload|null> {
    const webApi = (this.context as any).webAPI as ComponentFramework.WebApi;
    const anyApi: any = webApi;
    let resp:any;
    if (typeof anyApi.invokeUnboundFunction === "function") {
      resp = await anyApi.invokeUnboundFunction("vita_GetViewFilterSet", { EntityLogicalName: entity });
    } else {
      resp = await webApi.retrieveMultipleRecords(`vita_GetViewFilterSet(EntityLogicalName='${encodeURIComponent(entity)}')`, "");
    }
    const json = resp?.Response ?? resp?.entities?.[0]?.Response; if (!json) return null;
    try { return JSON.parse(json) as FilterPayload; } catch { return null; }
  }

  private findGridControl(): any|null {
    try { const pc:any[] = (window as any).Xrm?.App?.getVisibleControls?.() || []; const g = pc.find(c=>typeof c?.getFetchXml==="function" && typeof c?.setFilterXml==="function"); if (g) return g; } catch {}
    try { const controls:any[] = (window as any).Xrm?.Page?.ui?.controls?.get() || []; const g = controls.find(c=>typeof c?.getFetchXml==="function" && typeof c?.setFilterXml==="function"); if (g) return g; } catch {}
    return null;
  }

  private injectInCondition(fetchXml: string, field: string, values: string[], debug:boolean): string {
    const parser = new DOMParser(); const x = parser.parseFromString(fetchXml, "text/xml");
    const fetch = x.getElementsByTagName("fetch")[0]; if (!fetch) return fetchXml;
    const entity = fetch.getElementsByTagName("entity")[0]; if (!entity) return fetchXml;

    const filters = Array.from(entity.getElementsByTagName("filter")).filter(f=>f.parentElement===entity);
    let hostAnd = filters.find(f => (f.getAttribute("type")||"").toLowerCase()==="and");
    if (!hostAnd) { hostAnd = x.createElement("filter"); hostAnd.setAttribute("type","and"); filters.forEach(f=>{entity.removeChild(f); hostAnd!.appendChild(f);}); entity.appendChild(hostAnd); }

    const directConds = Array.from(hostAnd.children).filter(n=> n.nodeName==="condition" && (n as Element).getAttribute("attribute")?.toLowerCase()===field.toLowerCase()) as Element[];
    const existingIn = directConds.find(c => (c.getAttribute("operator")||"").toLowerCase()==="in");
    if (existingIn) {
      const existingVals = Array.from(existingIn.getElementsByTagName("value")).map(v=>v.textContent?.toLowerCase()||"");
      const want = [...new Set(values.map(v=>v.toLowerCase()))].sort();
      const have = [...new Set(existingVals)].sort();
      if (want.length===have.length && want.every((v,i)=>v===have[i])) { if (debug) this.log("identical IN found → skip"); return fetchXml; }
    }
    directConds.forEach(c=>hostAnd!.removeChild(c));

    hostAnd.appendChild(x.createComment(` VITA-INJECT field=${field} ts=${new Date().toISOString()} `));
    const cond = x.createElement("condition"); cond.setAttribute("attribute", field); cond.setAttribute("operator","in");
    values.forEach(g=>{ const v = x.createElement("value"); v.textContent = g; cond.appendChild(v); });
    hostAnd.appendChild(cond);

    return new XMLSerializer().serializeToString(x);
  }

  private log(m: string){ try{ console.log(`[VITA-PCF] ${m}`);}catch{} }
}
