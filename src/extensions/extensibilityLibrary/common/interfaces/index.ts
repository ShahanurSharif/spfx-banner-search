/**
 * Common Interfaces for Extensibility Library
 * 
 * Shared type definitions and interfaces used across the extensibility library
 */

// Re-export custom data source interfaces
export * from '../../components/dataSources/ICustomSearchDataSourceProps';

// Re-export layout interfaces
export * from '../../components/layouts/ListLayout';
export * from '../../components/layouts/CardLayout';

// Re-export filter interfaces
export * from '../../components/filters/CustomFilterComponent';

// Re-export service interfaces
export * from '../../services/SearchService';
export * from '../../services/GraphService';

/**
 * Common result item interface for normalized search results
 */
export interface ISearchResultItem {
    Title: string;
    Path: string;
    Summary?: string;
    Author?: string;
    Created?: string;
    Modified?: string;
    FileType?: string;
    FileExtension?: string;
    IsFromGraph?: boolean;
    GraphEntityType?: string;
    [key: string]: any;
}

/**
 * Common pagination interface
 */
export interface IPaginationInfo {
    currentPage: number;
    totalPages: number;
    pageSize: number;
    totalItems: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
}

/**
 * Common search context interface
 */
export interface ISearchContext {
    query: string;
    filters: IFilterContext[];
    sortBy: string;
    sortDirection: 'asc' | 'desc';
    pagination: IPaginationInfo;
    searchSources: string[];
}

/**
 * Filter context interface
 */
export interface IFilterContext {
    name: string;
    values: string[];
    operator: 'OR' | 'AND';
}

/**
 * Error handling interface
 */
export interface IExtensibilityError {
    code: string;
    message: string;
    details?: any;
    timestamp: Date;
    component: string;
}

/**
 * Configuration interface
 */
export interface IExtensibilityConfig {
    version: string;
    components: {
        dataSources: string[];
        layouts: string[];
        filters: string[];
    };
    settings: {
        cacheEnabled: boolean;
        cacheDuration: number;
        debugMode: boolean;
        telemetryEnabled: boolean;
    };
}