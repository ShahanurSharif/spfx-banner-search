/**
 * Interface for Custom Search Data Source Properties
 * 
 * Defines the configuration properties available for the custom search data source
 * These properties will be configurable through the property pane
 */

export interface ICustomSearchDataSourceProps {
    /**
     * KQL query template for SharePoint search
     * Default: '*' (all content)
     */
    queryTemplate?: string;

    /**
     * Comma-separated list of properties to retrieve from search results
     * Default: 'Title,Path,Author,Created,Modified,Summary'
     */
    selectProperties?: string;

    /**
     * Sort field for search results
     * Default: 'Rank' (relevance)
     */
    sortBy?: string;

    /**
     * Number of results to return per page
     * Default: 50, Max: 500
     */
    rowLimit?: number;

    /**
     * Enable Microsoft Graph search integration
     * Default: false
     */
    enableGraphSearch?: boolean;

    /**
     * Enable search refiners/filters
     * Default: true
     */
    enableRefiners?: boolean;

    /**
     * Custom refiners configuration (comma-separated)
     * Default: 'Author,FileType,ModifiedOOBDate'
     */
    customRefiners?: string;

    /**
     * Enable result trimming for duplicates
     * Default: true
     */
    trimDuplicates?: boolean;

    /**
     * Timeout for search requests (in milliseconds)
     * Default: 30000 (30 seconds)
     */
    timeout?: number;

    /**
     * Enable caching of search results
     * Default: true
     */
    enableCache?: boolean;

    /**
     * Cache duration in minutes
     * Default: 15
     */
    cacheDuration?: number;
}