/**
 * Interface definition for the Custom Extensibility Library
 * 
 * This interface extends the base PnP Modern Search extensibility interface
 * and defines additional methods specific to our custom implementation.
 */

import { IExtensibilityLibrary as IPnPExtensibilityLibrary } from '@pnp/modern-search-extensibility';

/**
 * Custom extensibility library interface
 * Extends the base PnP interface with additional functionality
 */
export interface IExtensibilityLibrary extends IPnPExtensibilityLibrary {
    /**
     * Initialize the extensibility library
     */
    initialize?(): void | Promise<void>;

    /**
     * Dispose of resources when the library is no longer needed
     */
    dispose?(): void;

    /**
     * Get the library version
     */
    getVersion?(): string;

    /**
     * Get library configuration
     */
    getConfiguration?(): Record<string, unknown>;
}