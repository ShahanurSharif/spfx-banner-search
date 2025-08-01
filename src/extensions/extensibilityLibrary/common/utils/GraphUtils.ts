/**
 * Microsoft Graph Utilities
 * 
 * Helper functions for Microsoft Graph API operations and data normalization
 */

/**
 * Graph API utility functions
 */
export class GraphUtils {

    /**
     * Normalize Microsoft Graph entity types to display-friendly names
     */
    public static getEntityDisplayName(entityType: string): string {
        const entityMappings: { [key: string]: string } = {
            'driveItem': 'File',
            'message': 'Email',
            'event': 'Calendar Event',
            'contact': 'Contact',
            'chatMessage': 'Teams Message',
            'site': 'SharePoint Site',
            'list': 'SharePoint List',
            'listItem': 'List Item',
            'person': 'Person',
            'group': 'Group',
            'team': 'Team',
            'channel': 'Channel'
        };

        return entityMappings[entityType.toLowerCase()] || entityType;
    }

    /**
     * Get appropriate icon for Graph entity type
     */
    public static getEntityIcon(entityType: string): string {
        const iconMappings: { [key: string]: string } = {
            'driveItem': 'OneDriveAdd',
            'message': 'Mail',
            'event': 'Calendar',
            'contact': 'Contact',
            'chatMessage': 'TeamsLogo',
            'site': 'SharePointLogo',
            'list': 'List',
            'listItem': 'ListMirrored',
            'person': 'Person',
            'group': 'Group',
            'team': 'TeamsLogo',
            'channel': 'TeamsLogo'
        };

        return iconMappings[entityType.toLowerCase()] || 'GenericScan';
    }

    /**
     * Extract file type from Graph drive item
     */
    public static getDriveItemType(driveItem: any): string {
        if (!driveItem) return 'Unknown';

        // Check if it's a folder
        if (driveItem.folder) {
            return 'Folder';
        }

        // Check file extension
        if (driveItem.file && driveItem.file.mimeType) {
            return this.getMimeTypeCategory(driveItem.file.mimeType);
        }

        // Fallback to name extension
        if (driveItem.name) {
            const extension = driveItem.name.split('.').pop()?.toLowerCase();
            if (extension) {
                return this.getFileExtensionCategory(extension);
            }
        }

        return 'File';
    }

    /**
     * Categorize by MIME type
     */
    private static getMimeTypeCategory(mimeType: string): string {
        if (!mimeType) return 'File';

        const type = mimeType.toLowerCase();

        if (type.startsWith('image/')) return 'Image';
        if (type.startsWith('video/')) return 'Video';
        if (type.startsWith('audio/')) return 'Audio';
        if (type.startsWith('text/')) return 'Text';

        // Specific document types
        if (type.includes('pdf')) return 'PDF';
        if (type.includes('word') || type.includes('document')) return 'Word Document';
        if (type.includes('excel') || type.includes('spreadsheet')) return 'Excel Spreadsheet';
        if (type.includes('powerpoint') || type.includes('presentation')) return 'PowerPoint';
        if (type.includes('zip') || type.includes('archive')) return 'Archive';

        return 'File';
    }

    /**
     * Categorize by file extension
     */
    private static getFileExtensionCategory(extension: string): string {
        const ext = extension.toLowerCase();

        // Document types
        if (['doc', 'docx'].includes(ext)) return 'Word Document';
        if (['xls', 'xlsx'].includes(ext)) return 'Excel Spreadsheet';
        if (['ppt', 'pptx'].includes(ext)) return 'PowerPoint';
        if (ext === 'pdf') return 'PDF';
        if (['txt', 'rtf'].includes(ext)) return 'Text';

        // Media types
        if (['jpg', 'jpeg', 'png', 'gif', 'bmp'].includes(ext)) return 'Image';
        if (['mp4', 'avi', 'mov', 'wmv'].includes(ext)) return 'Video';
        if (['mp3', 'wav', 'flac'].includes(ext)) return 'Audio';

        // Code types
        if (['js', 'ts', 'html', 'css', 'json'].includes(ext)) return 'Code';

        return 'File';
    }

    /**
     * Format Graph search request for optimal results
     */
    public static buildGraphSearchRequest(
        query: string,
        entityTypes: string[] = ['driveItem', 'message', 'event'],
        options: {
            top?: number;
            fields?: string[];
            enableQueryString?: boolean;
        } = {}
    ): any {
        const { top = 25, fields, enableQueryString = true } = options;

        const requests = entityTypes.map(entityType => {
            const request: any = {
                entityType,
                query: {
                    queryString: enableQueryString ? this.enhanceQueryString(query, entityType) : query
                },
                from: 0,
                size: Math.min(top, 25), // Graph API limit
                fields: fields || this.getDefaultFieldsForEntity(entityType)
            };

            // Add entity-specific configurations
            if (entityType === 'driveItem') {
                request.query.queryString += ' AND NOT(FileExtension:aspx)'; // Exclude SharePoint pages
            }

            return request;
        });

        return { requests };
    }

    /**
     * Enhance query string for specific entity types
     */
    private static enhanceQueryString(query: string, entityType: string): string {
        if (!query) return '*';

        // Add entity-specific query enhancements
        switch (entityType) {
            case 'driveItem':
                // Boost filename matches for files
                return `(${query}) OR filename:${query}`;
            
            case 'message':
                // Include subject and body searches for emails
                return `(${query}) OR subject:${query} OR body:${query}`;
            
            case 'event':
                // Include subject and location for events
                return `(${query}) OR subject:${query} OR location:${query}`;
            
            default:
                return query;
        }
    }

    /**
     * Get default fields for different entity types
     */
    private static getDefaultFieldsForEntity(entityType: string): string[] {
        const fieldMappings: { [key: string]: string[] } = {
            'driveItem': [
                'id', 'name', 'webUrl', 'size', 'createdDateTime', 'lastModifiedDateTime',
                'createdBy', 'lastModifiedBy', 'file', 'folder', 'parentReference'
            ],
            'message': [
                'id', 'subject', 'bodyPreview', 'from', 'toRecipients', 'receivedDateTime',
                'importance', 'isRead', 'hasAttachments', 'webLink'
            ],
            'event': [
                'id', 'subject', 'bodyPreview', 'start', 'end', 'location', 'organizer',
                'attendees', 'isAllDay', 'webLink', 'categories'
            ],
            'contact': [
                'id', 'displayName', 'emailAddresses', 'businessPhones', 'jobTitle',
                'companyName', 'department'
            ],
            'chatMessage': [
                'id', 'body', 'from', 'createdDateTime', 'importance', 'attachments'
            ]
        };

        return fieldMappings[entityType] || ['id', 'displayName'];
    }

    /**
     * Normalize Graph search response to common format
     */
    public static normalizeGraphResponse(graphResponse: any): any[] {
        const normalizedItems: any[] = [];

        if (!graphResponse.value || !Array.isArray(graphResponse.value)) {
            return normalizedItems;
        }

        for (const response of graphResponse.value) {
            if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
                for (const container of response.hitsContainers) {
                    if (container.hits && Array.isArray(container.hits)) {
                        for (const hit of container.hits) {
                            const normalizedItem = this.normalizeGraphItem(hit.resource, hit);
                            if (normalizedItem) {
                                normalizedItems.push(normalizedItem);
                            }
                        }
                    }
                }
            }
        }

        return normalizedItems;
    }

    /**
     * Normalize individual Graph item to common format
     */
    private static normalizeGraphItem(resource: any, hit: any): any {
        if (!resource) return null;

        const entityType = resource['@odata.type']?.replace('#microsoft.graph.', '') || 'unknown';
        
        return {
            // Common properties
            Title: this.extractTitle(resource, entityType),
            Path: resource.webUrl || resource.webLink || '#',
            Summary: this.extractSummary(resource, entityType),
            Author: this.extractAuthor(resource, entityType),
            Created: resource.createdDateTime || resource.dateTimeCreated,
            Modified: resource.lastModifiedDateTime || resource.dateTimeLastModified || resource.receivedDateTime,
            
            // Graph-specific properties
            IsFromGraph: true,
            GraphEntityType: entityType,
            GraphId: resource.id,
            HitId: hit.hitId,
            Rank: hit.rank,
            
            // Entity-specific properties
            ...this.extractEntitySpecificProperties(resource, entityType),
            
            // Original resource for advanced scenarios
            _originalResource: resource
        };
    }

    /**
     * Extract title based on entity type
     */
    private static extractTitle(resource: any, entityType: string): string {
        switch (entityType) {
            case 'driveItem':
                return resource.name || 'Untitled File';
            case 'message':
                return resource.subject || 'No Subject';
            case 'event':
                return resource.subject || 'No Title';
            case 'contact':
                return resource.displayName || 'Unknown Contact';
            default:
                return resource.displayName || resource.name || resource.title || 'Untitled';
        }
    }

    /**
     * Extract summary based on entity type
     */
    private static extractSummary(resource: any, entityType: string): string {
        switch (entityType) {
            case 'message':
                return resource.bodyPreview || '';
            case 'event':
                return resource.bodyPreview || resource.location?.displayName || '';
            case 'driveItem':
                return resource.file ? `${this.getDriveItemType(resource)} • ${this.formatFileSize(resource.size)}` : 'Folder';
            default:
                return resource.description || resource.bodyPreview || '';
        }
    }

    /**
     * Extract author based on entity type
     */
    private static extractAuthor(resource: any, entityType: string): string {
        switch (entityType) {
            case 'driveItem':
                return resource.createdBy?.user?.displayName || 'Unknown';
            case 'message':
                return resource.from?.emailAddress?.name || 'Unknown Sender';
            case 'event':
                return resource.organizer?.emailAddress?.name || 'Unknown Organizer';
            default:
                return resource.createdBy?.user?.displayName || 'Unknown';
        }
    }

    /**
     * Extract entity-specific properties
     */
    private static extractEntitySpecificProperties(resource: any, entityType: string): any {
        const properties: any = {};

        switch (entityType) {
            case 'driveItem':
                properties.FileType = this.getDriveItemType(resource);
                properties.FileSize = resource.size;
                properties.FileSizeFormatted = this.formatFileSize(resource.size);
                properties.ParentPath = resource.parentReference?.path;
                break;
                
            case 'message':
                properties.Importance = resource.importance;
                properties.IsRead = resource.isRead;
                properties.HasAttachments = resource.hasAttachments;
                properties.Recipients = resource.toRecipients?.map((r: any) => r.emailAddress?.name).join(', ');
                break;
                
            case 'event':
                properties.StartTime = resource.start?.dateTime;
                properties.EndTime = resource.end?.dateTime;
                properties.Location = resource.location?.displayName;
                properties.IsAllDay = resource.isAllDay;
                properties.Categories = resource.categories?.join(', ');
                break;
        }

        return properties;
    }

    /**
     * Format file size in human readable format
     */
    private static formatFileSize(bytes: number): string {
        if (!bytes || bytes === 0) return '0 B';
        
        const units = ['B', 'KB', 'MB', 'GB'];
        const unitIndex = Math.floor(Math.log(bytes) / Math.log(1024));
        const size = bytes / Math.pow(1024, unitIndex);
        
        return `${size.toFixed(1)} ${units[unitIndex]}`;
    }

    /**
     * Build Graph API URL with proper escaping
     */
    public static buildGraphUrl(endpoint: string, parameters: { [key: string]: any } = {}): string {
        const url = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
        
        const queryParams = new URLSearchParams();
        Object.keys(parameters).forEach(key => {
            if (parameters[key] !== undefined && parameters[key] !== null) {
                queryParams.append(key, parameters[key].toString());
            }
        });
        
        const queryString = queryParams.toString();
        return queryString ? `${url}?${queryString}` : url;
    }

    /**
     * Handle Graph API errors gracefully
     */
    public static handleGraphError(error: any): { message: string; code?: string; isRetryable: boolean } {
        if (!error) {
            return { message: 'Unknown error occurred', isRetryable: false };
        }

        // Handle Graph-specific error format
        if (error.code) {
            switch (error.code) {
                case 'Throttled':
                case 'TooManyRequests':
                    return { 
                        message: 'Request was throttled. Please try again later.', 
                        code: error.code,
                        isRetryable: true 
                    };
                    
                case 'Unauthorized':
                case 'Forbidden':
                    return { 
                        message: 'Access denied. Please check permissions.', 
                        code: error.code,
                        isRetryable: false 
                    };
                    
                case 'NotFound':
                    return { 
                        message: 'Requested resource was not found.', 
                        code: error.code,
                        isRetryable: false 
                    };
                    
                case 'ServiceUnavailable':
                    return { 
                        message: 'Microsoft Graph service is temporarily unavailable.', 
                        code: error.code,
                        isRetryable: true 
                    };
                    
                default:
                    return { 
                        message: error.message || 'Graph API error occurred', 
                        code: error.code,
                        isRetryable: false 
                    };
            }
        }

        // Handle standard errors
        return { 
            message: error.message || 'Unknown error occurred', 
            isRetryable: false 
        };
    }
}