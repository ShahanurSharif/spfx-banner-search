import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface SuggestionItem {
  id: string;
  suggestionTitle: string;
  suggestionSubtitle: string;
  path: string;
  fileType?: string;
  icon?: string;
}

export class SharePointSearchService {
  constructor(private context: WebPartContext, private siteUrl?: string, private debugSuggestions?: boolean) {}

  /**
   * Get static query suggestions from configuration
   */
  public getStaticSuggestions(staticSuggestions: string, query: string = ''): SuggestionItem[] {
    if (!staticSuggestions.trim()) return [];
    
    const suggestions = staticSuggestions
      .split(',')
      .map(s => s.trim())
      .filter(s => s.length > 0)
      .filter(s => query === '' || s.toLowerCase().includes(query.toLowerCase()))
      .slice(0, 10)
      .map((suggestion, index) => ({
        id: `static-${index}`,
        suggestionTitle: suggestion,
        suggestionSubtitle: 'Suggested search',
        path: '',
        fileType: ''
      } as SuggestionItem));

    return suggestions;
  }

  /**
   * Get zero-term suggestions (shown when search box is empty)
   */
  public getZeroTermSuggestions(zeroTermSuggestions: string): SuggestionItem[] {
    if (!zeroTermSuggestions.trim()) return [];
    
    const suggestions = zeroTermSuggestions
      .split(',')
      .map(s => s.trim())
      .filter(s => s.length > 0)
      .slice(0, 5)
      .map((suggestion, index) => ({
        id: `zero-term-${index}`,
        suggestionTitle: suggestion,
        suggestionSubtitle: 'Popular search',
        path: '',
        fileType: ''
      } as SuggestionItem));

    return suggestions;
  }

  public async fetchSuggestions(term: string, signal?: AbortSignal, limit: number = 10): Promise<SuggestionItem[]> {
    if (!term || !term.trim()) return [];
    const baseUrl = (this.siteUrl || this.context.pageContext.web.absoluteUrl).replace(/\/$/, "");
    
    // Use SharePoint Query API to get actual file results for typeahead (like the working version)
    const queryText = encodeURIComponent(`${term.trim()}*`);
    const selectProperties = encodeURIComponent("Title,Path,Author,LastModifiedTime,FileType,SiteName,SPWebUrl,HitHighlightedSummary,FileName,Name,FileLeafRef");
    
    // Match the exact format from the working version, using the provided limit
    const url = `${baseUrl}/_api/search/query?querytext='${queryText}'&selectproperties='${selectProperties}'&rowlimit=${limit}&trimduplicates=true`;
    
    if (this.debugSuggestions) {
      console.debug("[SharePoint Search API] URL:", url);
    }
    
    let resp;
    try {
      resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "3.0"
        }
      } as Record<string, unknown>);
    } catch (e: unknown) {
      if (this.debugSuggestions) console.debug("fetchSuggestions error", e);
      return [];
    }
    if (!resp.ok) {
      if (this.debugSuggestions) console.debug("fetchSuggestions not ok", url, resp.status);
      return [];
    }
    let json: unknown;
    try {
      json = await resp.json();
    } catch (e: unknown) {
      if (this.debugSuggestions) console.debug("fetchSuggestions json error", e);
      return [];
    }
    // Parse SharePoint Query API response to get actual file results
    let rows: unknown[] = [];
    const jsonData = json as Record<string, unknown>;
    if (jsonData?.d && typeof jsonData.d === 'object' && jsonData.d !== null) {
      const d = jsonData.d as Record<string, unknown>;
      if (d.query && typeof d.query === 'object' && d.query !== null) {
        const query = d.query as Record<string, unknown>;
        if (query?.PrimaryQueryResult && typeof query.PrimaryQueryResult === 'object' && query.PrimaryQueryResult !== null) {
          const primaryResult = query.PrimaryQueryResult as Record<string, unknown>;
          if (primaryResult?.RelevantResults && typeof primaryResult.RelevantResults === 'object' && primaryResult.RelevantResults !== null) {
            const relevantResults = primaryResult.RelevantResults as Record<string, unknown>;
            if (relevantResults?.Table && typeof relevantResults.Table === 'object' && relevantResults.Table !== null) {
              const table = relevantResults.Table as Record<string, unknown>;
              const tableRows = table.Rows;
              rows = Array.isArray(tableRows) ? tableRows : (Array.isArray((tableRows as Record<string, unknown>)?.results) ? (tableRows as Record<string, unknown>).results as unknown[] : []);
            }
          }
        }
      }
    } else if (jsonData?.PrimaryQueryResult && typeof jsonData.PrimaryQueryResult === 'object' && jsonData.PrimaryQueryResult !== null) {
      const primaryResult = jsonData.PrimaryQueryResult as Record<string, unknown>;
      if (primaryResult?.RelevantResults && typeof primaryResult.RelevantResults === 'object' && primaryResult.RelevantResults !== null) {
        const relevantResults = primaryResult.RelevantResults as Record<string, unknown>;
        if (relevantResults?.Table && typeof relevantResults.Table === 'object' && relevantResults.Table !== null) {
          const table = relevantResults.Table as Record<string, unknown>;
          const tableRows = table.Rows;
          rows = Array.isArray(tableRows) ? tableRows : [];
        }
      }
    }
    
    if (this.debugSuggestions) {
      console.debug("[SharePoint Search API] Rows found:", rows.length);
    }
    
    // Convert search results to SuggestionItem format with file names and metadata
    const suggestions = rows.map((r: unknown, index: number) => {
      try {
        const props: Record<string, unknown> = {};
        const row = r as Record<string, unknown>;
        
        // Handle both Cells and Cells.results structures
        const cellsData = row?.Cells as Record<string, unknown>;
        const cells = (cellsData?.results as unknown[]) || (row?.Cells as unknown[]) || [];
        cells.forEach((c: unknown) => {
          const cell = c as Record<string, unknown>;
          if (cell.Key && typeof cell.Key === 'string') {
            props[cell.Key] = cell.Value;
          }
        });
        
        // Extract file information
        const title = (props.FileName as string) || (props.Title as string) || (props.Name as string) || (props.FileLeafRef as string) || "(untitled)";
        const when = props.LastModifiedTime ? new Date(props.LastModifiedTime as string) : null;
        const subtitleParts = [
          (props.FileType as string) || "",
          (props.SiteName as string) || "",
          when ? when.toLocaleDateString() : ""
        ].filter(Boolean);
        
        return {
          id: (props.Path as string) || (props.UniqueId as string) || `suggestion-${index}-${Date.now()}`,
          suggestionTitle: title,
          suggestionSubtitle: subtitleParts.join(" · "),
          path: (props.Path as string) || "",
          fileType: (props.FileType as string) || ""
        } as SuggestionItem;
      } catch (error) {
        console.error(`[SharePoint Search API] Error processing row ${index}:`, error);
        return {
          id: `error-${index}`,
          suggestionTitle: `Error processing item ${index}`,
          suggestionSubtitle: "Error occurred",
          path: "",
          fileType: ""
        } as SuggestionItem;
      }
    });
    
    if (this.debugSuggestions) {
      console.debug("[SharePoint Search API] Returning", suggestions.length, "suggestions");
    }
    
    return suggestions;
  }
}
