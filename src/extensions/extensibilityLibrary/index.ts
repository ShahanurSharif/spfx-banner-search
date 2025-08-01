/**
 * My PnP Search Extensibility Library - Main Export
 * 
 * This is the main entry point for the extensibility library.
 * It exports the ExtensibilityLibrary class and registers it with PnP Modern Search.
 * 
 * To use this library:
 * 1. Build the SPFx solution
 * 2. Deploy to SharePoint App Catalog
 * 3. The library will be automatically discovered by PnP Modern Search web parts
 * 
 * @author SharePoint Framework Extensibility Library
 */

// Main extensibility library export
export { ExtensibilityLibrary } from './ExtensibilityLibrary';
export { IExtensibilityLibrary } from './IExtensibilityLibrary';

// Export service keys for registration
export {
    CustomSearchDataSourceServiceKey,
    ListLayoutServiceKey,
    CardLayoutServiceKey,
    SearchServiceKey,
    GraphServiceKey
} from './ExtensibilityLibrary';

// Export component classes
export { CustomSearchDataSource } from './components/dataSources/CustomSearchDataSource';
export { ListLayout } from './components/layouts/ListLayout';
export { CardLayout } from './components/layouts/CardLayout';
export { default as CustomFilterComponent } from './components/filters/CustomFilterComponent';

// Export services
export { SearchService } from './services/SearchService';
export { GraphService } from './services/GraphService';

// Export interfaces and types
export * from './common/interfaces';

// Export utilities
export { SearchUtils } from './common/utils/SearchUtils';
export { GraphUtils } from './common/utils/GraphUtils';

// Version information
export const EXTENSIBILITY_LIBRARY_VERSION = '1.0.0';
export const EXTENSIBILITY_LIBRARY_NAME = 'My PnP Search Extensibility Library';

/**
 * Library configuration object
 */
export const ExtensibilityLibraryConfig = {
    name: EXTENSIBILITY_LIBRARY_NAME,
    version: EXTENSIBILITY_LIBRARY_VERSION,
    description: 'Custom extensibility library for PnP Modern Search with SharePoint Search and Microsoft Graph integration',
    components: {
        dataSources: ['CustomSearchDataSource'],
        layouts: ['ListLayout', 'CardLayout'],
        filters: ['CustomFilterComponent']
    },
    features: [
        'SharePoint Search API integration',
        'Microsoft Graph search support',
        'Custom List and Card layouts',
        'Dynamic filter components',
        'Responsive design',
        'Accessibility support',
        'Handlebars template customization',
        'Caching support',
        'Error handling',
        'TypeScript support'
    ],
    dependencies: {
        '@pnp/modern-search-extensibility': '^1.5.0',
        '@microsoft/sp-webpart-base': '1.21.1',
        '@microsoft/sp-core-library': '1.21.1'
    },
    author: 'SharePoint Framework Development Team',
    license: 'MIT'
};

/**
 * Initialize the extensibility library
 * This function is called automatically when the library is loaded
 */
export function initializeExtensibilityLibrary(): void {
    console.log(`[ExtensibilityLibrary] Initializing ${EXTENSIBILITY_LIBRARY_NAME} v${EXTENSIBILITY_LIBRARY_VERSION}`);
    
    // Register any global configurations or event listeners here
    // This is a good place to set up telemetry, error handling, etc.
    
    // Log library features
    console.log('[ExtensibilityLibrary] Available features:', ExtensibilityLibraryConfig.features);
    console.log('[ExtensibilityLibrary] Available components:', ExtensibilityLibraryConfig.components);
    
    // Set up global error handling
    window.addEventListener('error', (event) => {
        if (event.error && event.error.stack && event.error.stack.includes('ExtensibilityLibrary')) {
            console.error('[ExtensibilityLibrary] Global error caught:', event.error);
        }
    });
    
    console.log('[ExtensibilityLibrary] Initialization completed successfully');
}

// Auto-initialize when the module is loaded
if (typeof window !== 'undefined') {
    // Delay initialization to ensure all dependencies are loaded
    setTimeout(initializeExtensibilityLibrary, 100);
}