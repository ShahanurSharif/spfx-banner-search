/**
 * Custom Search Data Source
 * 
 * Integrates SharePoint Search API with optional Microsoft Graph support
 * Provides flexible search capabilities across SharePoint and external data sources
 */

import {
    BaseDataSource,
    IDataSourceData,
    IDataContext,
    PagingBehavior,
    FilterBehavior,
    IDataFilter,
    ITemplateSlot
} from '@pnp/modern-search-extensibility';

import { IPropertyPaneGroup, PropertyPaneTextField, PropertyPaneToggle, PropertyPaneDropdown } from '@microsoft/sp-property-pane';


import { SearchService, ISearchResponse } from '../../services/SearchService';
import { GraphService, IGraphSearchResponse } from '../../services/GraphService';
import { ICustomSearchDataSourceProps } from './ICustomSearchDataSourceProps';

/**
 * Custom Search Data Source implementation
 * Extends BaseDataSource to provide SharePoint Search and Graph API integration
 */
export class CustomSearchDataSource extends BaseDataSource<ICustomSearchDataSourceProps> {
    
    private _searchService: SearchService;
    private _graphService: GraphService;
    private _itemCount: number = 0;

    constructor(serviceScope: any) { // eslint-disable-line @typescript-eslint/no-explicit-any
        super(serviceScope);
        
        // Initialize services
        this._searchService = new SearchService();
        this._graphService = new GraphService();
    }

    /**
     * Initialize the data source
     */
    public async onInit(): Promise<void> {
        try {
            // Initialize services with context
            await this._searchService.initialize(this.context as any); // eslint-disable-line @typescript-eslint/no-explicit-any
            await this._graphService.initialize(this.context as any); // eslint-disable-line @typescript-eslint/no-explicit-any
            
            console.log('[CustomSearchDataSource] Initialized successfully');
        } catch (error) {
            console.error('[CustomSearchDataSource] Initialization failed:', error);
        }
    }

    /**
     * Retrieve data from SharePoint Search and/or Microsoft Graph
     */
    public async getData(dataContext?: IDataContext): Promise<IDataSourceData> {
        try {
            console.log('[CustomSearchDataSource] Fetching data with context:', dataContext);

            const props = this.properties || {};
            const {
                queryTemplate = '*',
                enableGraphSearch = false,
                selectProperties = 'Title,Path,Author,Created,Modified,Summary',
                sortBy = 'Rank',
                enableRefiners = true,
                rowLimit = 50
            } = props;

            // Build search query
            let searchQuery = queryTemplate;
            
            // Apply query text from data context
            if (dataContext?.inputQueryText) {
                searchQuery = dataContext.inputQueryText;
            }

            // Apply filters if available
            if (dataContext?.filters && Array.isArray(dataContext.filters) && dataContext.filters.length > 0) {
                const filterQuery = this._buildFilterQuery(dataContext.filters);
                searchQuery = `${searchQuery} ${filterQuery}`;
            }

            // Calculate pagination
            const startRow = dataContext?.pageNumber ? (dataContext.pageNumber - 1) * rowLimit : 0;

            const searchResults: ISearchResponse = await this._searchService.search({
                queryText: searchQuery,
                selectProperties: selectProperties.split(','),
                sortList: [{ property: sortBy, direction: 0 }],
                rowLimit: rowLimit,
                startRow: startRow,
                refiners: enableRefiners ? 'Author,FileType,ModifiedOOBDate' : undefined
            });

            let graphResults: IGraphSearchResponse | null = null;



            // Execute Microsoft Graph search if enabled
            if (enableGraphSearch) {
                try {
                    graphResults = await this._graphService.search({
                        query: searchQuery,
                        entityTypes: ['driveItem', 'message', 'event'],
                        top: Math.min(rowLimit, 25) // Graph API has limits
                    });

                    // Merge results if both sources return data
                    if (graphResults && graphResults.items && graphResults.items.length > 0) {
                        const normalizedGraphResults = this._normalizeGraphResults(graphResults.items);
                        searchResults.items = [...searchResults.items, ...normalizedGraphResults];
                    }
                } catch (graphError) {
                    console.warn('[CustomSearchDataSource] Graph search failed, continuing with SharePoint results:', graphError);
                }
            }

            // Update item count for pagination
            this._itemCount = searchResults.totalRows || searchResults.items.length;

            // Prepare result data
            const resultData: IDataSourceData = {
                items: searchResults.items || [],
                totalRows: this._itemCount,
                filters: searchResults.refinementResults || [],
                // Add custom properties for template context
                searchQuery: searchQuery,
                hasGraphResults: enableGraphSearch && graphResults && graphResults.items.length > 0,
                resultSources: {
                    sharePoint: searchResults.items?.length || 0,
                    graph: graphResults?.items?.length || 0
                }
            };

            console.log('[CustomSearchDataSource] Data fetched successfully:', {
                itemCount: resultData.items.length,
                totalRows: resultData.totalRows,
                hasFilters: (resultData.filters || []).length > 0
            });

            return resultData;

        } catch (error) {
            console.error('[CustomSearchDataSource] Error fetching data:', error);
            
            // Return empty result set on error
            return {
                items: [],
                totalRows: 0,
                filters: [],
                error: error.message || 'Unknown error occurred'
            };
        }
    }

    /**
     * Get property pane configuration for the data source
     */
    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
        return [
            {
                groupName: 'Search Configuration',
                groupFields: [
                    PropertyPaneTextField('queryTemplate', {
                        label: 'Query Template',
                        description: 'KQL query template (* for all content)',
                        value: this.properties?.queryTemplate || '*'
                    }),
                    PropertyPaneTextField('selectProperties', {
                        label: 'Select Properties',
                        description: 'Comma-separated list of properties to retrieve',
                        value: this.properties?.selectProperties || 'Title,Path,Author,Created,Modified,Summary'
                    }),
                    PropertyPaneDropdown('sortBy', {
                        label: 'Sort By',
                        options: [
                            { key: 'Rank', text: 'Relevance' },
                            { key: 'Created', text: 'Created Date' },
                            { key: 'Modified', text: 'Modified Date' },
                            { key: 'Title', text: 'Title' }
                        ],
                        selectedKey: this.properties?.sortBy || 'Rank'
                    }),
                    PropertyPaneTextField('rowLimit', {
                        label: 'Results Per Page',
                        description: 'Number of results to display per page (1-500)',
                        value: (this.properties?.rowLimit || 50).toString()
                    })
                ]
            },
            {
                groupName: 'Advanced Options',
                groupFields: [
                    PropertyPaneToggle('enableGraphSearch', {
                        label: 'Enable Microsoft Graph Search',
                        checked: this.properties?.enableGraphSearch || false
                    }),
                    PropertyPaneToggle('enableRefiners', {
                        label: 'Enable Refiners',
                        checked: this.properties?.enableRefiners !== false
                    })
                ]
            }
        ];
    }

    /**
     * Get paging behavior - supports pagination
     */
    public getPagingBehavior(): PagingBehavior {
        return PagingBehavior.Dynamic;
    }

    /**
     * Get filter behavior - supports dynamic filtering
     */
    public getFilterBehavior(): FilterBehavior {
        return FilterBehavior.Dynamic;
    }

    /**
     * Get applied filters (used for static filtering)
     */
    public getAppliedFilters(): IDataFilter[] {
        return []; // Dynamic filtering - no static filters applied by data source
    }

    /**
     * Get total item count for pagination
     */
    public getItemCount(): number {
        return this._itemCount;
    }

    /**
     * Get available template slots for this data source
     */
    public getTemplateSlots(): ITemplateSlot[] {
        return [
            {
                slotName: 'Title',
                slotField: 'Title'
            },
            {
                slotName: 'Summary',
                slotField: 'Summary'
            },
            {
                slotName: 'Author',
                slotField: 'Author'
            },
            {
                slotName: 'Modified',
                slotField: 'Modified'
            },
            {
                slotName: 'Path',
                slotField: 'Path'
            }
        ];
    }

    /**
     * Handle property updates
     */
    public onPropertyUpdate(propertyPath: string, oldValue: unknown, newValue: unknown): void {
        console.log(`[CustomSearchDataSource] Property '${propertyPath}' changed from '${oldValue}' to '${newValue}'`);
        
        // Convert string values to appropriate types
        if (propertyPath === 'rowLimit') {
            const numValue = parseInt(newValue as string, 10);
            if (!isNaN(numValue) && numValue > 0 && numValue <= 500) {
                this.properties.rowLimit = numValue;
            }
        }
        
        // Re-render when properties change
        if (this.render) {
            const renderResult = this.render();
            if (renderResult instanceof Promise) {
                renderResult.catch(error => console.error('Render error:', error));
            }
        }
    }

    /**
     * Get sortable fields for the data source
     */
    public getSortableFields(): string[] {
        return ['Rank', 'Title', 'Created', 'Modified', 'Author'];
    }

    /**
     * Build KQL filter query from applied filters
     */
    private _buildFilterQuery(filters: IDataFilter[]): string {
        const filterParts: string[] = [];

        for (const filter of filters) {
            if (filter.values && filter.values.length > 0) {
                const filterValues = filter.values
                    .filter((v: any) => v.selected) // eslint-disable-line @typescript-eslint/no-explicit-any
                    .map((v: any) => `${filter.filterName}:"${v.value}"`) // eslint-disable-line @typescript-eslint/no-explicit-any
                    .join(' OR ');
                
                if (filterValues) {
                    filterParts.push(`(${filterValues})`);
                }
            }
        }

        return filterParts.length > 0 ? `AND (${filterParts.join(' AND ')})` : '';
    }

    /**
     * Normalize Microsoft Graph results to match SharePoint Search format
     */
    private _normalizeGraphResults(graphItems: unknown[]): unknown[] {
        return graphItems.map((item: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
            Title: item.name || item.subject || 'Untitled',
            Path: item.webUrl || item.webLink || '#',
            Author: item.createdBy?.user?.displayName || item.from?.emailAddress?.name || 'Unknown',
            Created: item.createdDateTime || item.dateTimeCreated,
            Modified: item.lastModifiedDateTime || item.dateTimeLastModified,
            Summary: item.summary || item.body?.content?.substring(0, 200) || '',
            FileType: this._getGraphItemType(item),
            IsFromGraph: true,
            GraphEntityType: item['@odata.type'] || 'unknown'
        }));
    }

    /**
     * Determine file type from Graph item
     */
    private _getGraphItemType(item: any): string { // eslint-disable-line @typescript-eslint/no-explicit-any
        if (item['@odata.type']) {
            const type = item['@odata.type'].toLowerCase();
            if (type.includes('message')) return 'Email';
            if (type.includes('event')) return 'Event';
            if (type.includes('driveitem')) return item.file?.mimeType || 'File';
        }
        return 'Unknown';
    }
}