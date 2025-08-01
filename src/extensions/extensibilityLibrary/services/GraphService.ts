/**
 * Microsoft Graph Service
 * 
 * Provides Microsoft Graph API integration for cross-platform search
 * Supports OneDrive, Outlook, Teams, and other Microsoft 365 services
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';

/**
 * Graph search request parameters interface
 */
export interface IGraphSearchRequest {
    query: string;
    entityTypes: string[]; // driveItem, message, event, etc.
    top?: number;
    from?: number;
    fields?: string[];
}

/**
 * Graph search response interface
 */
export interface IGraphSearchResponse {
    items: any[];
    totalItems: number;
    moreResultsAvailable: boolean;
}

/**
 * Graph Service implementation
 * Provides Microsoft Graph API functionality
 */
export class GraphService {
    private _graphClient: MSGraphClient;
    private _cache: Map<string, { data: IGraphSearchResponse; timestamp: number }> = new Map();
    private _cacheTimeout: number = 10 * 60 * 1000; // 10 minutes

    /**
     * Initialize the Graph service with SharePoint context
     */
    public async initialize(context: WebPartContext): Promise<void> {
        try {
            this._graphClient = await context.msGraphClientFactory.getClient('3') as any; // eslint-disable-line @typescript-eslint/no-explicit-any
            
            console.log('[GraphService] Initialized successfully');
        } catch (error) {
            console.error('[GraphService] Initialization failed:', error);
            throw error;
        }
    }

    /**
     * Execute Microsoft Graph search query
     */
    public async search(request: IGraphSearchRequest): Promise<IGraphSearchResponse> {
        try {
            console.log('[GraphService] Executing Graph search:', request);

            // Check cache first
            const cacheKey = this._generateCacheKey(request);
            const cachedResult = this._getFromCache(cacheKey);
            if (cachedResult) {
                console.log('[GraphService] Returning cached Graph result');
                return cachedResult;
            }

            // Build Graph search request
            const searchBody = this._buildGraphSearchBody(request);

            // Execute Graph search
            const response = await this._graphClient
                .api('/search/query')
                .post(searchBody);

            const searchResult = this._parseGraphResponse(response);

            // Cache the result
            this._addToCache(cacheKey, searchResult);

            console.log('[GraphService] Graph search completed successfully:', {
                totalItems: searchResult.totalItems,
                itemCount: searchResult.items.length
            });

            return searchResult;

        } catch (error) {
            console.error('[GraphService] Graph search failed:', error);
            
            // Return empty result on error rather than throwing
            return {
                items: [],
                totalItems: 0,
                moreResultsAvailable: false
            };
        }
    }

    /**
     * Get user's OneDrive files
     */
    public async getOneDriveFiles(query?: string, top: number = 25): Promise<any[]> {
        try {
            let endpoint = '/me/drive/root/search(q=\'{query}\')';
            
            if (query) {
                endpoint = endpoint.replace('{query}', encodeURIComponent(query));
            } else {
                // If no query, get recent files
                endpoint = '/me/drive/recent';
            }

            const response = await this._graphClient
                .api(endpoint)
                .top(top)
                .select('id,name,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,size,file')
                .get();

            return response.value || [];

        } catch (error) {
            console.error('[GraphService] OneDrive files request failed:', error);
            return [];
        }
    }

    /**
     * Get user's recent emails
     */
    public async getRecentEmails(query?: string, top: number = 25): Promise<any[]> {
        try {
            const endpoint = '/me/messages';
            const params: any = {
                top: top,
                select: 'id,subject,bodyPreview,from,toRecipients,receivedDateTime,webLink,hasAttachments'
            };

            if (query) {
                params.search = `"${query}"`;
            }

            let apiCall = this._graphClient.api(endpoint);
            
            // Apply parameters
            Object.keys(params).forEach(key => {
                if (key === 'select') {
                    apiCall = apiCall.select(params[key]);
                } else if (key === 'top') {
                    apiCall = apiCall.top(params[key]);
                } else if (key === 'search') {
                    // For search, we'll use query parameters instead
                    apiCall = apiCall.query({ '$search': params[key] });
                }
            });

            const response = await apiCall.get();
            return response.value || [];

        } catch (error) {
            console.error('[GraphService] Recent emails request failed:', error);
            return [];
        }
    }

    /**
     * Get user's calendar events
     */
    public async getCalendarEvents(query?: string, top: number = 25): Promise<any[]> {
        try {
            const endpoint = '/me/events';
            const params: any = {
                top: top,
                select: 'id,subject,bodyPreview,start,end,location,organizer,webLink,attendees'
            };

            if (query) {
                params.search = `"${query}"`;
            }

            let apiCall = this._graphClient.api(endpoint);
            
            // Apply parameters
            Object.keys(params).forEach(key => {
                if (key === 'select') {
                    apiCall = apiCall.select(params[key]);
                } else if (key === 'top') {
                    apiCall = apiCall.top(params[key]);
                } else if (key === 'search') {
                    // For search, we'll use query parameters instead
                    apiCall = apiCall.query({ '$search': params[key] });
                }
            });

            const response = await apiCall.get();
            return response.value || [];

        } catch (error) {
            console.error('[GraphService] Calendar events request failed:', error);
            return [];
        }
    }

    /**
     * Get user profile information
     */
    public async getUserProfile(): Promise<any> {
        try {
            const response = await this._graphClient
                .api('/me')
                .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation')
                .get();

            return response;

        } catch (error) {
            console.error('[GraphService] Get user profile failed:', error);
            return null;
        }
    }

    /**
     * Clear Graph cache
     */
    public clearCache(): void {
        this._cache.clear();
        console.log('[GraphService] Cache cleared');
    }

    /**
     * Build Microsoft Graph search request body
     */
    private _buildGraphSearchBody(request: IGraphSearchRequest): any {
        const searchRequests = request.entityTypes.map(entityType => ({
            entityType: entityType,
            query: {
                queryString: request.query
            },
            from: request.from || 0,
            size: Math.min(request.top || 25, 25), // Graph API limits
            fields: request.fields || this._getDefaultFieldsForEntityType(entityType)
        }));

        return {
            requests: searchRequests
        };
    }

    /**
     * Get default fields for entity type
     */
    private _getDefaultFieldsForEntityType(entityType: string): string[] {
        switch (entityType.toLowerCase()) {
            case 'driveitem':
                return ['id', 'name', 'webUrl', 'createdDateTime', 'lastModifiedDateTime', 'createdBy', 'size'];
            case 'message':
                return ['id', 'subject', 'from', 'receivedDateTime', 'webLink', 'bodyPreview'];
            case 'event':
                return ['id', 'subject', 'start', 'end', 'location', 'organizer', 'webLink'];
            default:
                return ['id', 'displayName', 'webUrl'];
        }
    }

    /**
     * Parse Microsoft Graph search response
     */
    private _parseGraphResponse(response: any): IGraphSearchResponse {
        const result: IGraphSearchResponse = {
            items: [],
            totalItems: 0,
            moreResultsAvailable: false
        };

        try {
            if (response.value && Array.isArray(response.value)) {
                for (const searchResponse of response.value) {
                    if (searchResponse.hitsContainers && Array.isArray(searchResponse.hitsContainers)) {
                        for (const hitsContainer of searchResponse.hitsContainers) {
                            if (hitsContainer.hits && Array.isArray(hitsContainer.hits)) {
                                // Extract items from hits
                                const items = hitsContainer.hits.map((hit: any) => ({
                                    ...hit.resource,
                                    hitId: hit.hitId,
                                    rank: hit.rank,
                                    summary: hit.summary
                                }));
                                
                                result.items.push(...items);
                                result.totalItems += hitsContainer.total || items.length;
                                result.moreResultsAvailable = result.moreResultsAvailable || hitsContainer.moreResultsAvailable;
                            }
                        }
                    }
                }
            }

        } catch (error) {
            console.error('[GraphService] Error parsing Graph response:', error);
        }

        return result;
    }

    /**
     * Generate cache key from search request
     */
    private _generateCacheKey(request: IGraphSearchRequest): string {
        return `graph_${btoa(JSON.stringify(request))}`;
    }

    /**
     * Get result from cache if not expired
     */
    private _getFromCache(key: string): IGraphSearchResponse | null {
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
    private _addToCache(key: string, data: IGraphSearchResponse): void {
        this._cache.set(key, {
            data,
            timestamp: Date.now()
        });
    }
}