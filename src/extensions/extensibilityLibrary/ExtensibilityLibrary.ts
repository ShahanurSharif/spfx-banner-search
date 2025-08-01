/**
 * My PnP Search Extensibility Library
 * 
 * This library provides custom extensions for PnP Modern Search web parts:
 * - Custom SharePoint Search and Microsoft Graph data sources
 * - List and Card layout renderers
 * - Custom filter components for dynamic refiners
 * - Modular services (SearchService, GraphService)
 * 
 * @author SharePoint Framework Extensibility Library
 */

import { 
    IExtensibilityLibrary, 
    ILayoutDefinition,
    IComponentDefinition,
    ISuggestionProviderDefinition,
    IQueryModifierDefinition,
    IDataSourceDefinition,
    IAdaptiveCardAction
} from '@pnp/modern-search-extensibility';

import { ServiceKey } from '@microsoft/sp-core-library';
import * as Handlebars from 'handlebars';

// Import custom components
import { CustomSearchDataSource } from './components/dataSources/CustomSearchDataSource';
import { ListLayout } from './components/layouts/ListLayout';
import { CardLayout } from './components/layouts/CardLayout';

// Import service keys
import { SearchService } from './services/SearchService';
import { GraphService } from './services/GraphService';

// Service keys
export const CustomSearchDataSourceServiceKey: ServiceKey<any> = 
    ServiceKey.create<any>('MyExtensibilityLibrary:CustomSearchDataSource', CustomSearchDataSource);

export const ListLayoutServiceKey: ServiceKey<any> = 
    ServiceKey.create<any>('MyExtensibilityLibrary:ListLayout', ListLayout);

export const CardLayoutServiceKey: ServiceKey<any> = 
    ServiceKey.create<any>('MyExtensibilityLibrary:CardLayout', CardLayout);

export const SearchServiceKey: ServiceKey<SearchService> = 
    ServiceKey.create<SearchService>('MyExtensibilityLibrary:SearchService', SearchService);

export const GraphServiceKey: ServiceKey<GraphService> = 
    ServiceKey.create<GraphService>('MyExtensibilityLibrary:GraphService', GraphService);

/**
 * Main extensibility library implementation
 * Registers all custom components with PnP Modern Search
 */
export class ExtensibilityLibrary implements IExtensibilityLibrary {

    /**
     * Returns custom data sources
     */
    public getCustomDataSources(): IDataSourceDefinition[] {
        return [
            {
                name: 'Custom SharePoint Search',
                key: 'CustomSearchDataSource',
                iconName: 'SearchAndApps',
                serviceKey: CustomSearchDataSourceServiceKey as any // eslint-disable-line @typescript-eslint/no-explicit-any
            }
        ];
    }

    /**
     * Returns custom layouts for results rendering
     */
    public getCustomLayouts(): ILayoutDefinition[] {
        // Import templates as strings
        const listLayoutTemplate = require('./components/layouts/templates/ListLayout.html').default || require('./components/layouts/templates/ListLayout.html'); // eslint-disable-line @typescript-eslint/no-var-requires
        const cardLayoutTemplate = require('./components/layouts/templates/CardLayout.html').default || require('./components/layouts/templates/CardLayout.html'); // eslint-disable-line @typescript-eslint/no-var-requires

        return [
            {
                name: 'Custom List Layout',
                key: 'CustomListLayout',
                type: 'ResultsLayout' as any,
                iconName: 'List',
                templateContent: listLayoutTemplate,
                renderType: 'Handlebars' as any,
                serviceKey: ListLayoutServiceKey as any // eslint-disable-line @typescript-eslint/no-explicit-any
            },
            {
                name: 'Custom Card Layout',
                key: 'CustomCardLayout',
                type: 'ResultsLayout' as any,
                iconName: 'GridViewMedium',
                templateContent: cardLayoutTemplate,
                renderType: 'Handlebars' as any,
                serviceKey: CardLayoutServiceKey as any // eslint-disable-line @typescript-eslint/no-explicit-any
            }
        ];
    }

    /**
     * Returns custom web components (not implemented in this version)
     */
    public getCustomWebComponents(): IComponentDefinition<any>[] {
        return [];
    }

    /**
     * Returns custom suggestion providers (not implemented in this version)
     */
    public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
        return [];
    }

    /**
     * Returns custom query modifiers (not implemented in this version)
     */
    public getCustomQueryModifiers(): IQueryModifierDefinition[] {
        return [];
    }

    /**
     * Register custom Handlebars helpers and partials
     */
    public registerHandlebarsCustomizations(handlebarsNamespace: typeof Handlebars): void {
        // Register custom helper for formatting dates
        handlebarsNamespace.registerHelper('formatDate', (dateString: string) => {
            if (!dateString) return '';
            const date = new Date(dateString);
            return date.toLocaleDateString();
        });

        // Register custom helper for highlighting search terms
        handlebarsNamespace.registerHelper('highlight', (text: string, searchTerm: string) => {
            if (!text || !searchTerm) return text;
            // Escape special regex characters to prevent injection
            const escapedTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const regex = new RegExp(`(${escapedTerm})`, 'gi');
            return text.replace(regex, '<mark>$1</mark>');
        });

        // Register custom helper for truncating text
        handlebarsNamespace.registerHelper('truncate', (text: string, length: number = 100) => {
            if (!text) return '';
            return text.length > length ? text.substring(0, length) + '...' : text;
        });

        // Register custom helper for checking if value exists
        handlebarsNamespace.registerHelper('ifExists', function(value: any, options: any) {
            if (value && value !== '') {
                return options.fn(this);
            }
            return options.inverse(this);
        });

        console.log('[ExtensibilityLibrary] Custom Handlebars helpers registered');
    }

    /**
     * Handle adaptive card actions (placeholder implementation)
     */
    public invokeCardAction(action: IAdaptiveCardAction): void {
        console.log('[ExtensibilityLibrary] Adaptive card action invoked:', action);
        
        // Handle different action types
        switch (action.type) {
            case 'Action.OpenUrl':
                if (action.url) {
                    window.open(action.url, '_blank');
                }
                break;
            case 'Action.Submit':
                // Handle form submission
                console.log('[ExtensibilityLibrary] Form submitted with data:', action.data);
                break;
            default:
                console.warn('[ExtensibilityLibrary] Unknown action type:', action.type);
        }
    }
}