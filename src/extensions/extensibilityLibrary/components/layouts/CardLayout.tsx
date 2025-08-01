/**
 * Custom Card Layout Component
 * 
 * Provides a modern card-based view for search results
 * Supports responsive grid layout, thumbnails, and rich metadata
 */

import { BaseLayout } from '@pnp/modern-search-extensibility';
import { IPropertyPaneField, PropertyPaneToggle, PropertyPaneDropdown, PropertyPaneSlider } from '@microsoft/sp-property-pane';

/**
 * Card layout properties interface
 */
export interface ICardLayoutProps {
    cardsPerRow?: number;
    showThumbnails?: boolean;
    showAuthor?: boolean;
    showDate?: boolean;
    cardHeight?: number;
    enableAnimations?: boolean;
    shadowDepth?: string;
    borderRadius?: number;
}

/**
 * Custom Card Layout implementation
 * Extends BaseLayout to provide card-style rendering
 */
export class CardLayout extends BaseLayout<ICardLayoutProps> {

    constructor(serviceScope: any) { // eslint-disable-line @typescript-eslint/no-explicit-any
        super(serviceScope);
    }

    /**
     * Initialize the layout
     */
    public async onInit(): Promise<void> {
        console.log('[CardLayout] Initialized');
    }

    /**
     * Get property pane fields for layout configuration
     */
    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneSlider('cardsPerRow', {
                label: 'Cards Per Row',
                min: 1,
                max: 6,
                value: this.properties?.cardsPerRow || 3,
                showValue: true,
                step: 1
            }),
            PropertyPaneSlider('cardHeight', {
                label: 'Card Height (px)',
                min: 200,
                max: 500,
                value: this.properties?.cardHeight || 300,
                showValue: true,
                step: 25
            }),
            PropertyPaneSlider('borderRadius', {
                label: 'Border Radius (px)',
                min: 0,
                max: 20,
                value: this.properties?.borderRadius || 8,
                showValue: true,
                step: 2
            }),
            PropertyPaneToggle('showThumbnails', {
                label: 'Show Thumbnails',
                checked: this.properties?.showThumbnails !== false
            }),
            PropertyPaneToggle('showAuthor', {
                label: 'Show Author',
                checked: this.properties?.showAuthor !== false
            }),
            PropertyPaneToggle('showDate', {
                label: 'Show Date',
                checked: this.properties?.showDate !== false
            }),
            PropertyPaneToggle('enableAnimations', {
                label: 'Enable Animations',
                checked: this.properties?.enableAnimations !== false
            }),
            PropertyPaneDropdown('shadowDepth', {
                label: 'Shadow Depth',
                options: [
                    { key: 'none', text: 'None' },
                    { key: 'light', text: 'Light' },
                    { key: 'medium', text: 'Medium' },
                    { key: 'deep', text: 'Deep' }
                ],
                selectedKey: this.properties?.shadowDepth || 'light'
            })
        ];
    }

    /**
     * Handle property updates
     */
    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        console.log(`[CardLayout] Property '${propertyPath}' changed from '${oldValue}' to '${newValue}'`);
        
        // Additional property validation can be added here
        if (propertyPath === 'cardsPerRow') {
            const numValue = parseInt(newValue, 10);
            if (!isNaN(numValue) && numValue > 0 && numValue <= 6) {
                this.properties.cardsPerRow = numValue;
            }
        }
        
        if (propertyPath === 'cardHeight') {
            const numValue = parseInt(newValue, 10);
            if (!isNaN(numValue) && numValue >= 200 && numValue <= 500) {
                this.properties.cardHeight = numValue;
            }
        }
    }
}