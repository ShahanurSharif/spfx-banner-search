/**
 * Custom List Layout Component
 * 
 * Provides a clean, accessible list view for search results
 * Supports sorting, filtering, and responsive design
 */

import { BaseLayout } from '@pnp/modern-search-extensibility';
import { IPropertyPaneField, PropertyPaneToggle, PropertyPaneDropdown } from '@microsoft/sp-property-pane';

/**
 * List layout properties interface
 */
export interface IListLayoutProps {
    showMetadata?: boolean;
    showThumbnails?: boolean;
    compactMode?: boolean;
    sortBy?: string;
    maxItems?: number;
    enableHover?: boolean;
}

/**
 * Custom List Layout implementation
 * Extends BaseLayout to provide list-style rendering
 */
export class ListLayout extends BaseLayout<IListLayoutProps> {

    constructor(serviceScope: any) { // eslint-disable-line @typescript-eslint/no-explicit-any
        super(serviceScope);
    }

    /**
     * Initialize the layout
     */
    public async onInit(): Promise<void> {
        console.log('[ListLayout] Initialized');
    }

    /**
     * Get property pane fields for layout configuration
     */
    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneToggle('showMetadata', {
                label: 'Show Metadata',
                checked: this.properties?.showMetadata !== false
            }),
            PropertyPaneToggle('showThumbnails', {
                label: 'Show Thumbnails',
                checked: this.properties?.showThumbnails !== false
            }),
            PropertyPaneToggle('compactMode', {
                label: 'Compact Mode',
                checked: this.properties?.compactMode || false
            }),
            PropertyPaneToggle('enableHover', {
                label: 'Enable Hover Effects',
                checked: this.properties?.enableHover !== false
            }),
            PropertyPaneDropdown('sortBy', {
                label: 'Default Sort Order',
                options: [
                    { key: 'Rank', text: 'Relevance' },
                    { key: 'Title', text: 'Title' },
                    { key: 'Created', text: 'Date Created' },
                    { key: 'Modified', text: 'Date Modified' },
                    { key: 'Author', text: 'Author' }
                ],
                selectedKey: this.properties?.sortBy || 'Rank'
            })
        ];
    }

    /**
     * Handle property updates
     */
    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        console.log(`[ListLayout] Property '${propertyPath}' changed from '${oldValue}' to '${newValue}'`);
        
        // Additional property validation can be added here
        if (propertyPath === 'maxItems') {
            const numValue = parseInt(newValue, 10);
            if (!isNaN(numValue) && numValue > 0) {
                this.properties.maxItems = Math.min(numValue, 1000);
            }
        }
    }
}