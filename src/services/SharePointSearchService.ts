import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface SuggestionItem {
  id: string;
  suggestionTitle: string;
  suggestionSubtitle: string;
  path: string;
  icon?: string;
}

export class SharePointSearchService {
  constructor(private context: WebPartContext, private siteUrl?: string, private debugSuggestions?: boolean) {}

  public async fetchSuggestions(term: string, signal?: AbortSignal): Promise<SuggestionItem[]> {
    if (!term || !term.trim()) return [];
    const baseUrl = (this.siteUrl || this.context.pageContext.web.absoluteUrl).replace(/\/$/, "");
    const qp = new URLSearchParams();
    qp.set("querytext", `'${term.trim()}*'`);
    qp.set("selectproperties", "Title,Path,Author,LastModifiedTime,FileType,SiteName,SPWebUrl,HitHighlightedSummary,FileName,Name,FileLeafRef");
    qp.set("rowlimit", "10");
    qp.set("trimduplicates", "true");
    const url = `${baseUrl}/_api/search/query?${qp.toString()}`;
    let resp;
    try {
      resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "3.0"
        }
      } as any);
    } catch (e) {
      if (this.debugSuggestions) console.debug("fetchSuggestions error", e);
      return [];
    }
    if (!resp.ok) {
      if (this.debugSuggestions) console.debug("fetchSuggestions not ok", url, resp.status);
      return [];
    }
    let json: any;
    try {
      json = await resp.json();
    } catch (e) {
      if (this.debugSuggestions) console.debug("fetchSuggestions json error", e);
      return [];
    }
    // Try both d.query and direct response shapes
    let rows = [];
    if (json?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows) {
      const tableRows = json.d.query.PrimaryQueryResult.RelevantResults.Table.Rows;
      rows = Array.isArray(tableRows) ? tableRows : (tableRows.results || []);
    } else if (json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows) {
      rows = json.PrimaryQueryResult.RelevantResults.Table.Rows;
    }
    if (this.debugSuggestions) {
      const firstRow = rows[0]?.Cells;
      console.debug("[Typeahead]", url, "rows:", rows.length, firstRow);
    }
    return rows.map((r: any) => {
      const props: Record<string, any> = {};
      (r?.Cells || []).forEach((c: any) => (props[c.Key] = c.Value));
      const title = props.FileName || props.Title || props.Name || props.FileLeafRef || "(untitled)";
      const when = props.LastModifiedTime ? new Date(props.LastModifiedTime) : null;
      const subtitleParts = [
        props.FileType || "",
        props.SiteName || "",
        when ? when.toLocaleDateString() : ""
      ].filter(Boolean);
      return {
        id: props.Path || props.UniqueId || (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : String(Math.random())),
        suggestionTitle: title,
        suggestionSubtitle: subtitleParts.join(" · "),
        path: props.Path || ""
      } as SuggestionItem;
    });
  }
}
