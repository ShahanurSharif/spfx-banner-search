/**
 * Search Utilities
 * 
 * Helper functions for search operations, query building, and result processing
 */

/**
 * Build KQL query from various search parameters
 */
export class SearchUtils {

    /**
     * Escape special characters in KQL query
     */
    public static escapeKQLValue(value: string): string {
        if (!value) return '';
        
        // Escape special KQL characters
        return value.replace(/[(){}[\]"':]/g, '\\$&');
    }

    /**
     * Build KQL property filter
     */
    public static buildPropertyFilter(property: string, values: string[], operator: 'OR' | 'AND' = 'OR'): string {
        if (!values || values.length === 0) return '';
        
        const escapedValues = values.map(v => `"${this.escapeKQLValue(v)}"`);
        const joinOperator = operator === 'OR' ? ' OR ' : ' AND ';
        
        if (escapedValues.length === 1) {
            return `${property}:${escapedValues[0]}`;
        }
        
        return `(${escapedValues.map(v => `${property}:${v}`).join(joinOperator)})`;
    }

    /**
     * Build date range filter for KQL
     */
    public static buildDateRangeFilter(property: string, startDate?: Date, endDate?: Date): string {
        if (!startDate && !endDate) return '';
        
        const formatDate = (date: Date): string => {
            return date.toISOString().split('T')[0];
        };

        if (startDate && endDate) {
            return `${property}:${formatDate(startDate)}..${formatDate(endDate)}`;
        } else if (startDate) {
            return `${property}>=${formatDate(startDate)}`;
        } else if (endDate) {
            return `${property}<=${formatDate(endDate)}`;
        }
        
        return '';
    }

    /**
     * Build wildcard query for partial matches
     */
    public static buildWildcardQuery(property: string, value: string): string {
        if (!value) return '';
        
        const escapedValue = this.escapeKQLValue(value);
        return `${property}:*${escapedValue}*`;
    }

    /**
     * Combine multiple KQL filters with AND operator
     */
    public static combineFilters(filters: string[]): string {
        const validFilters = filters.filter(f => f && f.trim());
        if (validFilters.length === 0) return '';
        
        return validFilters.join(' AND ');
    }

    /**
     * Extract file extension from path
     */
    public static getFileExtension(path: string): string {
        if (!path) return '';
        
        const lastDot = path.lastIndexOf('.');
        const lastSlash = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
        
        if (lastDot > lastSlash && lastDot > 0) {
            return path.substring(lastDot + 1).toLowerCase();
        }
        
        return '';
    }

    /**
     * Determine file type category from extension
     */
    public static getFileTypeCategory(extension: string): string {
        if (!extension) return 'Unknown';
        
        const ext = extension.toLowerCase();
        
        // Document types
        if (['doc', 'docx', 'pdf', 'txt', 'rtf', 'odt'].includes(ext)) {
            return 'Document';
        }
        
        // Spreadsheet types
        if (['xls', 'xlsx', 'csv', 'ods'].includes(ext)) {
            return 'Spreadsheet';
        }
        
        // Presentation types
        if (['ppt', 'pptx', 'odp'].includes(ext)) {
            return 'Presentation';
        }
        
        // Image types
        if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) {
            return 'Image';
        }
        
        // Video types
        if (['mp4', 'avi', 'mov', 'wmv', 'flv', 'webm', 'mkv'].includes(ext)) {
            return 'Video';
        }
        
        // Audio types
        if (['mp3', 'wav', 'flac', 'aac', 'ogg', 'wma'].includes(ext)) {
            return 'Audio';
        }
        
        // Code types
        if (['js', 'ts', 'html', 'css', 'json', 'xml', 'sql', 'cs', 'java', 'py', 'php'].includes(ext)) {
            return 'Code';
        }
        
        // Archive types
        if (['zip', 'rar', '7z', 'tar', 'gz'].includes(ext)) {
            return 'Archive';
        }
        
        return 'File';
    }

    /**
     * Highlight search terms in text
     */
    public static highlightSearchTerms(text: string, searchTerms: string[], className: string = 'highlight'): string {
        if (!text || !searchTerms || searchTerms.length === 0) return text;
        
        let highlightedText = text;
        
        for (const term of searchTerms) {
            if (term && term.length > 2) { // Only highlight terms longer than 2 characters
                const escapedTerm = this.escapeRegExp(term);
                const regex = new RegExp(`(${escapedTerm})`, 'gi');
                highlightedText = highlightedText.replace(regex, `<span class="${className}">$1</span>`);
            }
        }
        
        return highlightedText;
    }

    /**
     * Escape special regex characters
     */
    private static escapeRegExp(text: string): string {
        return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Truncate text to specified length with ellipsis
     */
    public static truncateText(text: string, maxLength: number = 100, suffix: string = '...'): string {
        if (!text) return '';
        if (text.length <= maxLength) return text;
        
        return text.substring(0, maxLength - suffix.length) + suffix;
    }

    /**
     * Clean and normalize search query
     */
    public static normalizeQuery(query: string): string {
        if (!query) return '';
        
        // Remove extra whitespace and normalize
        return query.trim().replace(/\s+/g, ' ');
    }

    /**
     * Extract search terms from query string
     */
    public static extractSearchTerms(query: string): string[] {
        if (!query) return [];
        
        const normalizedQuery = this.normalizeQuery(query);
        
        // Split by spaces but respect quoted phrases
        const terms: string[] = [];
        const regex = /"([^"]+)"|(\S+)/g;
        let match;
        
        while ((match = regex.exec(normalizedQuery)) !== null) {
            terms.push(match[1] || match[2]);
        }
        
        return terms.filter(term => term && term.length > 0);
    }

    /**
     * Format file size in human readable format
     */
    public static formatFileSize(bytes: number): string {
        if (!bytes || bytes === 0) return '0 B';
        
        const units = ['B', 'KB', 'MB', 'GB', 'TB'];
        const unitIndex = Math.floor(Math.log(bytes) / Math.log(1024));
        const size = bytes / Math.pow(1024, unitIndex);
        
        return `${size.toFixed(1)} ${units[unitIndex]}`;
    }

    /**
     * Format date in user-friendly format
     */
    public static formatDate(dateString: string | Date, includeTime: boolean = false): string {
        if (!dateString) return '';
        
        const date = typeof dateString === 'string' ? new Date(dateString) : dateString;
        if (isNaN(date.getTime())) return '';
        
        const now = new Date();
        const diffMs = now.getTime() - date.getTime();
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        
        // Show relative dates for recent items
        if (diffDays === 0) {
            return includeTime ? `Today ${date.toLocaleTimeString()}` : 'Today';
        } else if (diffDays === 1) {
            return includeTime ? `Yesterday ${date.toLocaleTimeString()}` : 'Yesterday';
        } else if (diffDays < 7) {
            return includeTime ? `${diffDays} days ago` : `${diffDays} days ago`;
        }
        
        // Use standard date format for older items
        const options: Intl.DateTimeFormatOptions = {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        };
        
        if (includeTime) {
            options.hour = '2-digit';
            options.minute = '2-digit';
        }
        
        return date.toLocaleDateString(undefined, options);
    }

    /**
     * Parse SharePoint managed property names to display names
     */
    public static formatPropertyName(propertyName: string): string {
        if (!propertyName) return '';
        
        // Common SharePoint managed property mappings
        const propertyMappings: { [key: string]: string } = {
            'ModifiedOOBDate': 'Modified Date',
            'CreatedOOBDate': 'Created Date',
            'LastModifiedTime': 'Last Modified',
            'AuthorOWSUSER': 'Author',
            'EditorOWSUSER': 'Editor',
            'FileExtension': 'File Type',
            'ContentTypeId': 'Content Type',
            'SPSiteURL': 'Site',
            'SiteName': 'Site Name',
            'ListId': 'List',
            'DocumentLinkOWSMTXT': 'Document Link'
        };
        
        // Return mapped name if available
        if (propertyMappings[propertyName]) {
            return propertyMappings[propertyName];
        }
        
        // Convert camelCase/PascalCase to readable format
        return propertyName
            .replace(/([A-Z])/g, ' $1')
            .replace(/^./, str => str.toUpperCase())
            .trim();
    }

    /**
     * Build search query suggestions based on partial input
     */
    public static generateQuerySuggestions(partialQuery: string, commonTerms: string[] = []): string[] {
        if (!partialQuery || partialQuery.length < 2) return [];
        
        const suggestions: string[] = [];
        const lowerQuery = partialQuery.toLowerCase();
        
        // Add common terms that start with the query
        for (const term of commonTerms) {
            if (term.toLowerCase().startsWith(lowerQuery)) {
                suggestions.push(term);
            }
        }
        
        // Add wildcard suggestions
        suggestions.push(`${partialQuery}*`);
        suggestions.push(`*${partialQuery}*`);
        
        // Add property-specific suggestions
        const properties = ['title', 'author', 'filename', 'content'];
        for (const prop of properties) {
            suggestions.push(`${prop}:${partialQuery}`);
        }
        
        return suggestions.slice(0, 8); // Limit to 8 suggestions
    }
}