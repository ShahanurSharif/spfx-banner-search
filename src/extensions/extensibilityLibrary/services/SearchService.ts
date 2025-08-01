/**
 * SharePoint Search Service
 * 
 * Provides comprehensive SharePoint Search API integration
 * Handles KQL queries, refiners, sorting, and pagination
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Search request parameters interface
 */
export interface ISearchRequest {
    queryText: string;
    selectProperties?: string[];
    refiners?: string;
    sortList?: ISortField[];
    rowLimit?: number;
    startRow?: number;
    trimDuplicates?: boolean;
    enableQueryRules?: boolean;
    enableStemming?: boolean;
    enableNicknames?: boolean;
    enablePhonetic?: boolean;
}

/**
 * Sort field interface
 */
export interface ISortField {
    property: string;
    direction: number; // 0 = ascending, 1 = descending
}

/**
 * Search response interface
 */
export interface ISearchResponse {
    items: any[];
    totalRows: number;
    refinementResults?: any[];
    queryRuleId?: string;
    spellingSuggestion?: string;
}

/**
 * Search Service implementation
 * Provides SharePoint Search API functionality
 */
export class SearchService {
    private _context: WebPartContext;
    private _searchEndpoint: string;
    private _cache: Map<string, { data: ISearchResponse; timestamp: number }> = new Map();
    private _cacheTimeout: number = 15 * 60 * 1000; // 15 minutes

    /**
     * Initialize the search service with SharePoint context
     */
    public async initialize(context: WebPartContext): Promise<void> {
        this._context = context;
        this._searchEndpoint = `${context.pageContext.web.absoluteUrl}/_api/search/query`;
        
        console.log('[SearchService] Initialized with endpoint:', this._searchEndpoint);
    }

    /**
     * Execute SharePoint search query
     */
    public async search(request: ISearchRequest): Promise<ISearchResponse> {
        try {
            console.log('[SearchService] Executing search:', request);

            // Check cache first
            const cacheKey = this._generateCacheKey(request);
            const cachedResult = this._getFromCache(cacheKey);
            if (cachedResult) {
                console.log('[SearchService] Returning cached result');
                return cachedResult;
            }

            // Build search query
            const searchQuery = this._buildSearchQuery(request);
            const postBody = JSON.stringify({
                request: searchQuery
            });

            // Execute search request
            const response: SPHttpClientResponse = await this._context.spHttpClient.post(
                this._searchEndpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;charset=utf-8',
                        'odata-version': ''
                    },
                    body: postBody
                }
            );

            if (!response.ok) {
                throw new Error(`Search request failed: ${response.status} ${response.statusText}`);
            }

            const responseData = await response.json();
            const searchResult = this._parseSearchResponse(responseData);

            // Cache the result
            this._addToCache(cacheKey, searchResult);

            console.log('[SearchService] Search completed successfully:', {
                totalRows: searchResult.totalRows,
                itemCount: searchResult.items.length
            });

            return searchResult;

        } catch (error) {
            console.error('[SearchService] Search failed:', error);
            throw error;
        }
    }

    /**
     * Get search suggestions based on query text
     */
    public async getSuggestions(queryText: string, numberOfSuggestions: number = 5): Promise<string[]> {
        try {
            const suggestEndpoint = `${this._context.pageContext.web.absoluteUrl}/_api/search/suggest`;
            const suggestQuery = {
                querytext: queryText,
                count: numberOfSuggestions,
                fPreQuerySuggestions: true,
                fPostQuerySuggestions: true
            };

            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                `${suggestEndpoint}?${this._buildQueryString(suggestQuery)}`,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Suggestions request failed: ${response.status} ${response.statusText}`);
            }

            const responseData = await response.json();
            const suggestions: string[] = [];

            // Parse suggestions from response
            if (responseData.d && responseData.d.Suggest && responseData.d.Suggest.Queries) {
                responseData.d.Suggest.Queries.forEach((query: any) => {
                    if (query.Query) {
                        suggestions.push(query.Query);
                    }
                });
            }

            return suggestions;

        } catch (error) {
            console.error('[SearchService] Get suggestions failed:', error);
            return [];
        }
    }

    /**
     * Clear search cache
     */
    public clearCache(): void {
        this._cache.clear();
        console.log('[SearchService] Cache cleared');
    }

    /**
     * Build SharePoint search query from request parameters
     */
    private _buildSearchQuery(request: ISearchRequest): any {
        const query: any = {
            Querytext: request.queryText,
            RowLimit: request.rowLimit || 50,
            StartRow: request.startRow || 0,
            TrimDuplicates: request.trimDuplicates !== false,
            EnableQueryRules: request.enableQueryRules !== false,
            EnableStemming: request.enableStemming !== false,
            EnableNicknames: request.enableNicknames !== false,
            EnablePhonetic: request.enablePhonetic !== false
        };

        // Add select properties
        if (request.selectProperties && request.selectProperties.length > 0) {
            query.SelectProperties = {
                results: request.selectProperties
            };
        }

        // Add refiners
        if (request.refiners) {
            query.Refiners = request.refiners;
        }

        // Add sort list
        if (request.sortList && request.sortList.length > 0) {
            query.SortList = {
                results: request.sortList.map(sort => ({
                    Property: sort.property,
                    Direction: sort.direction
                }))
            };
        }

        return query;
    }

    /**
     * Parse SharePoint search response
     */
    private _parseSearchResponse(response: any): ISearchResponse {
        const result: ISearchResponse = {
            items: [],
            totalRows: 0
        };

        try {
            if (response.d && response.d.query && response.d.query.PrimaryQueryResult) {
                const primaryResult = response.d.query.PrimaryQueryResult;

                // Extract search results
                if (primaryResult.RelevantResults) {
                    result.totalRows = primaryResult.RelevantResults.TotalRows || 0;
                    
                    if (primaryResult.RelevantResults.Table && primaryResult.RelevantResults.Table.Rows) {
                        result.items = this._parseSearchItems(primaryResult.RelevantResults.Table.Rows.results);
                    }
                }

                // Extract refinement results
                if (primaryResult.RefinementResults && primaryResult.RefinementResults.Refiners) {
                    result.refinementResults = this._parseRefinementResults(
                        primaryResult.RefinementResults.Refiners.results
                    );
                }

                // Extract other metadata
                result.queryRuleId = response.d.query.Properties?.results?.find(
                    (p: any) => p.Key === 'QueryRuleId'
                )?.Value;

                result.spellingSuggestion = response.d.query.SpellingSuggestion;
            }

        } catch (error) {
            console.error('[SearchService] Error parsing search response:', error);
        }

        return result;
    }

    /**
     * Parse search result items from SharePoint response
     */
    private _parseSearchItems(rows: any[]): any[] {
        const items: any[] = [];

        for (const row of rows) {
            const item: any = {};
            
            if (row.Cells && row.Cells.results) {
                for (const cell of row.Cells.results) {
                    item[cell.Key] = cell.Value;
                }
            }

            items.push(item);
        }

        return items;
    }

    /**
     * Parse refinement results from SharePoint response
     */
    private _parseRefinementResults(refiners: any[]): any[] {
        const refinementResults: any[] = [];

        for (const refiner of refiners) {
            const refinement: any = {
                name: refiner.Name,
                values: []
            };

            if (refiner.Entries && refiner.Entries.results) {
                refinement.values = refiner.Entries.results.map((entry: any) => ({
                    value: entry.RefinementValue,
                    count: entry.RefinementCount
                }));
            }

            refinementResults.push(refinement);
        }

        return refinementResults;
    }

    /**
     * Generate cache key from search request
     */
    private _generateCacheKey(request: ISearchRequest): string {
        return btoa(JSON.stringify(request));
    }

    /**
     * Get result from cache if not expired
     */
    private _getFromCache(key: string): ISearchResponse | null {
        const cached = this._cache.get(key);
        if (cached && (Date.now() - cached.timestamp) < this._cacheTimeout) {
            return cached.data;
        }
        
        // Remove expired cache entry
        if (cached) {
            this._cache.delete(key);
        }
        
        return null;
    }

    /**
     * Add result to cache
     */
    private _addToCache(key: string, data: ISearchResponse): void {
        this._cache.set(key, {
            data,
            timestamp: Date.now()
        });
    }

    /**
     * Build query string from parameters
     */
    private _buildQueryString(params: any): string {
        return Object.keys(params)
            .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
            .join('&');
    }
}