import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISemanticColors } from '@fluentui/react/lib/Styling';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
// Import package.json for version information
const packageInfo = require('../../../package.json');

// Interface for extensibility libraries
export interface ExtensibilityLibrary {
  id: string;
  name: string;
  purpose: string;
  manifestGuid: string;
  enabled: boolean;
}

import * as strings from 'SpfxBannerSearchWebPartStrings';
import SpfxBannerSearch from './components/SpfxBannerSearch';
import { ISpfxBannerSearchProps } from './components/ISpfxBannerSearchProps';

export interface ISpfxBannerSearchWebPartProps {
  // Banner & Search Box configuration
  gradientStartColor: string;
  gradientEndColor: string;
  showCircleAnimation: boolean;
  minHeight: number;
  titleFontSize: number;
  bannerTitleColor: string;
  bannerTitle: string;
  searchBoxPlaceholder: string;
  searchBoxBorderRadius: number;
  searchBoxHeight: number;
  
  // Search behavior configuration
  queryTemplate: string;
  resultsWebPartId: string;
  enableSuggestions: boolean;
  suggestionsLimit: number;
  openingBehavior: string;
  
  // Query suggestions configuration
  enableQuerySuggestions: boolean;
  staticSuggestions: string;
  enableZeroTermSuggestions: boolean;
  zeroTermSuggestions: string;
  suggestionsProvider: string;
  
  // Custom search suggestions configuration
  hubSiteId: string;
  imageRelativeUrl: string;
  
  // Dynamic data source configuration
  useDynamicDataSource: boolean;
  dynamicDataSourceId: string;
  pageEnvironmentProperty: string;
  siteProperty: string;
  userProperty: string;
  queryStringProperty: string;
  searchProperty: string;
  
  // About and settings configuration - panels now handled natively
  
  // Extensibility libraries configuration
  extensibilityLibraries: string; // JSON string of ExtensibilityLibrary[]
  
  // Redirect configuration
  redirectToSearchPage: boolean;
  searchPageUrl: string;
  searchMethod: string;
  searchParameterName: string;
}

export default class SpfxBannerSearchWebPart extends BaseClientSideWebPart<ISpfxBannerSearchWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _semanticColors: Partial<ISemanticColors> = {};

  public render(): void {
    const element: React.ReactElement<ISpfxBannerSearchProps> = React.createElement(
      SpfxBannerSearch,
      {
        gradientStartColor: this.properties.gradientStartColor || '#0078d4',
        gradientEndColor: this.properties.gradientEndColor || '#106ebe',
        showCircleAnimation: this.properties.showCircleAnimation !== false,
        minHeight: this.properties.minHeight || 400,
        titleFontSize: this.properties.titleFontSize || 48,
        bannerTitleColor: this.properties.bannerTitleColor || '#ffffff',
        bannerTitle: this._processDynamicTitle(this.properties.bannerTitle || 'Find What You Need'),
        searchBoxPlaceholder: this.properties.searchBoxPlaceholder || 'Search everything...',
        searchBoxBorderRadius: this.properties.searchBoxBorderRadius || 4,
        searchBoxHeight: this.properties.searchBoxHeight || 32,
        queryTemplate: this.properties.queryTemplate || '*',
        enableSuggestions: this.properties.enableSuggestions !== false,
        suggestionsLimit: this.properties.suggestionsLimit || 10,
        openingBehavior: this.properties.openingBehavior || 'new-tab',
        enableQuerySuggestions: this.properties.enableQuerySuggestions !== false,
        staticSuggestions: this.properties.staticSuggestions || '',
        enableZeroTermSuggestions: this.properties.enableZeroTermSuggestions !== false,
        zeroTermSuggestions: this.properties.zeroTermSuggestions || '',
        suggestionsProvider: this.properties.suggestionsProvider || 'static',
        hubSiteId: this.properties.hubSiteId || '',
        imageRelativeUrl: this.properties.imageRelativeUrl || '',
        useDynamicDataSource: this.properties.useDynamicDataSource !== false,
        dynamicDataSourceId: this.properties.dynamicDataSourceId || '',
        pageEnvironmentProperty: this.properties.pageEnvironmentProperty || '',
        siteProperty: this.properties.siteProperty || '',
        userProperty: this.properties.userProperty || '',
        queryStringProperty: this.properties.queryStringProperty || '',
        searchProperty: this.properties.searchProperty || '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        semanticColors: this._semanticColors,
        context: this.context,
        onSearchQuery: this._onSearchQuery.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Process dynamic title with user property replacements
   * Supports placeholders: {email}, {firstname}, {lastname}, {displayname}
   */
  private _processDynamicTitle(title: string): string {
    if (!title || title.indexOf('{') === -1) {
      return title;
    }

    const user = this.context.pageContext.user;
    let processedTitle = title;

    // Replace user property placeholders
    const replacements = {
      '{email}': user.email || '',
      '{firstname}': user.displayName ? user.displayName.split(' ')[0] : '',
      '{lastname}': user.displayName ? user.displayName.split(' ').slice(1).join(' ') : '',
      '{displayname}': user.displayName || '',
      '{loginname}': user.loginName || ''
    };

    // Apply replacements using safer string replacement approach
    Object.keys(replacements).forEach(placeholder => {
      // Use simple string replacement instead of regex to avoid security concerns
      while (processedTitle.toLowerCase().includes(placeholder.toLowerCase())) {
        const index = processedTitle.toLowerCase().indexOf(placeholder.toLowerCase());
        if (index !== -1) {
          processedTitle = processedTitle.substring(0, index) + 
                          replacements[placeholder] + 
                          processedTitle.substring(index + placeholder.length);
        } else {
          break;
        }
      }
    });

    return processedTitle;
  }

  /**
   * Handle search query submission - publish as dynamic data
   */
  private _onSearchQuery(queryText: string): void {
    this._currentSearchQuery = queryText;
    
    // Option 1: Redirect to search results page
    if (this.properties.redirectToSearchPage && this.properties.searchPageUrl) {
      const method = this.properties.searchMethod || 'query-string';
      let searchUrl: string;
      
      if (method === 'url-fragment') {
        // URL fragment: /page#searchterm
        searchUrl = `${this.properties.searchPageUrl}#${encodeURIComponent(queryText)}`;
      } else {
        // Query string parameter: /page?param=searchterm
        const paramName = this.properties.searchParameterName || 'q';
        searchUrl = `${this.properties.searchPageUrl}?${paramName}=${encodeURIComponent(queryText)}`;
      }
      
      window.location.href = searchUrl;
      return;
    }
    
    // Option 2: Default behavior - Dynamic Data (same page)
    this.context.dynamicDataSourceManager.notifyPropertyChanged('inputQueryText');
  }

  /**
   * Get dynamic data source properties - allows other web parts to connect to this search box
   */
  public getPropertyDefinitions(): ReadonlyArray<{ id: string; title: string }> {
    return [
      {
        id: 'inputQueryText',
        title: 'Search Query Text'
      }
    ];
  }

  /**
   * Get the current search query value for dynamic data consumers
   */
  public getPropertyValue(propertyId: string): string {
    switch (propertyId) {
      case 'inputQueryText':
        return this._currentSearchQuery || '';
      default:
        throw new Error('Bad property id');
    }
  }

  private _currentSearchQuery: string = '';

  protected onInit(): Promise<void> {
    // Register as dynamic data source
    this.context.dynamicDataSourceManager.initializeSource(this);
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this._semanticColors = semanticColors;
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Native SharePoint Panel methods
  private _showExtensibilityPanel(): void {
    this._createExtensibilityPanel();
  }

  // Get extensibility libraries from properties or create defaults
  private _getExtensibilityLibraries(): ExtensibilityLibrary[] {
    try {
      if (this.properties.extensibilityLibraries) {
        return JSON.parse(this.properties.extensibilityLibraries);
      }
    } catch (error) {
      console.warn('Failed to parse extensibilityLibraries:', error);
    }
    
    // Return default libraries
    return [
      {
        id: '1',
        name: 'PnP Modern Search Extensibility',
        purpose: 'Custom result layouts and refiners',
        manifestGuid: '4588a19c-6c21-4f42-9bb6-9a7a3b62b1fa',
        enabled: true
      },
      {
        id: '2',
        name: 'Advanced Query Builders',
        purpose: 'Enhanced search query construction',
        manifestGuid: 'b2c3d4e5-f6g7-8901-bcde-f23456789012',
        enabled: false
      }
    ];
  }

  // Save extensibility libraries to properties
  private _saveExtensibilityLibraries(libraries: ExtensibilityLibrary[]): void {
    this.properties.extensibilityLibraries = JSON.stringify(libraries);
    this.context.propertyPane.refresh();
  }

  // HTML escape function for security
  private _escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // Setup global actions for extensibility panel
  private _setupExtensibilityActions(): void {
    const self = this;
    
    (window as any).spfxExtensibilityActions = {
      toggleLibrary: (index: number, enabled: boolean) => {
        const libraries = self._getExtensibilityLibraries();
        if (libraries[index]) {
          libraries[index].enabled = enabled;
          self._currentExtensibilityLibraries = libraries;
        }
      },
      
      updateLibraryName: (index: number, name: string) => {
        const libraries = self._getExtensibilityLibraries();
        if (libraries[index]) {
          libraries[index].name = name;
          self._currentExtensibilityLibraries = libraries;
        }
      },
      
      updateLibraryPurpose: (index: number, purpose: string) => {
        const libraries = self._getExtensibilityLibraries();
        if (libraries[index]) {
          libraries[index].purpose = purpose;
          self._currentExtensibilityLibraries = libraries;
        }
      },
      
      updateLibraryGuid: (index: number, guid: string) => {
        const libraries = self._getExtensibilityLibraries();
        if (libraries[index]) {
          libraries[index].manifestGuid = guid;
          self._currentExtensibilityLibraries = libraries;
        }
      },
      
      removeLibrary: (index: number) => {
        const libraries = self._currentExtensibilityLibraries || self._getExtensibilityLibraries();
        libraries.splice(index, 1);
        self._currentExtensibilityLibraries = libraries;
        self._updateLibrariesContainer();
      },
      
      addNewLibrary: () => {
        const libraries = self._currentExtensibilityLibraries || self._getExtensibilityLibraries();
        const newId = (Math.max(...libraries.map(lib => parseInt(lib.id) || 0)) + 1).toString();
        libraries.push({
          id: newId,
          name: '',
          purpose: '',
          manifestGuid: '',
          enabled: false
        });
        self._currentExtensibilityLibraries = libraries;
        self._updateLibrariesContainer();
      },
      
      saveLibraries: () => {
        const libraries = self._currentExtensibilityLibraries || self._getExtensibilityLibraries();
        self._saveExtensibilityLibraries(libraries);
        self._currentExtensibilityLibraries = null;
        alert('Extensibility libraries configuration saved successfully!');
        document.getElementById('extensibility-panel-overlay')?.remove();
      },
      
      resetToDefaults: () => {
        if (confirm('Are you sure you want to reset to default libraries? This will remove all custom configurations.')) {
          self._currentExtensibilityLibraries = [
            {
              id: '1',
              name: 'PnP Modern Search Extensibility',
              purpose: 'Custom result layouts and refiners',
              manifestGuid: '4588a19c-6c21-4f42-9bb6-9a7a3b62b1fa',
              enabled: true
            },
            {
              id: '2',
              name: 'Advanced Query Builders',
              purpose: 'Enhanced search query construction',
              manifestGuid: 'b2c3d4e5-f6g7-8901-bcde-f23456789012',
              enabled: false
            }
          ];
          self._updateLibrariesContainer();
        }
      }
    };
  }

  // Temporary storage for unsaved changes
  private _currentExtensibilityLibraries: ExtensibilityLibrary[] | null = null;

  // Update libraries container without flickering
  private _updateLibrariesContainer(): void {
    const container = document.getElementById('libraries-container');
    if (container) {
      const libraries = this._currentExtensibilityLibraries || this._getExtensibilityLibraries();
      container.innerHTML = this._generateLibrariesHTML(libraries);
    }
  }

  // Generate libraries HTML with Fluent UI components
  private _generateLibrariesHTML(libraries: ExtensibilityLibrary[]): string {
    return libraries.map((lib, index) => `
      <div style="margin-bottom: 20px; padding: 16px; border: 1px solid #edebe9; border-radius: 4px; background: #faf9f8;">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;">
          <h4 style="margin: 0; color: #323130; font-size: 16px; font-weight: 600; font-family: 'Segoe UI', sans-serif;">Library ${index + 1}</h4>
          <label style="display: flex; align-items: center; cursor: pointer;">
            <input type="checkbox" id="lib-enabled-${index}" ${lib.enabled ? 'checked' : ''} style="margin-right: 8px;" onchange="window.spfxExtensibilityActions.toggleLibrary(${index}, this.checked);">
            <span style="font-size: 14px; color: #323130; font-family: 'Segoe UI', sans-serif;">Enabled</span>
          </label>
        </div>
        
        <!-- Name/Purpose Field with Fluent UI styling -->
        <div style="margin-bottom: 12px;">
          <label class="ms-Label" style="display: block; margin-bottom: 4px; font-size: 14px; font-weight: 600; color: #323130; font-family: 'Segoe UI', sans-serif;">Name/Purpose:</label>
          <div class="ms-TextField">
            <input type="text" id="lib-name-${index}" value="${this._escapeHtml(lib.name)}" placeholder="Enter library name and purpose" 
                   class="ms-TextField-field" 
                   style="width: calc(100% - 24px); padding: 8px 12px; border: 1px solid #8a8886; border-radius: 2px; font-size: 14px; font-family: 'Segoe UI', sans-serif; background: #ffffff;" 
                   onchange="window.spfxExtensibilityActions.updateLibraryName(${index}, this.value);"
                   onfocus="this.style.borderColor='#0078d4'; this.style.boxShadow='0 0 0 1px #0078d4';"
                   onblur="this.style.borderColor='#8a8886'; this.style.boxShadow='none';">
          </div>
        </div>
        
        <!-- Purpose Description Field with Fluent UI styling -->
        <div style="margin-bottom: 12px;">
          <label class="ms-Label" style="display: block; margin-bottom: 4px; font-size: 14px; font-weight: 600; color: #323130; font-family: 'Segoe UI', sans-serif;">Purpose Description:</label>
          <div class="ms-TextField">
            <input type="text" id="lib-purpose-${index}" value="${this._escapeHtml(lib.purpose)}" placeholder="Describe what this library does" 
                   class="ms-TextField-field"
                   style="width: calc(100% - 24px); padding: 8px 12px; border: 1px solid #8a8886; border-radius: 2px; font-size: 14px; font-family: 'Segoe UI', sans-serif; background: #ffffff;" 
                   onchange="window.spfxExtensibilityActions.updateLibraryPurpose(${index}, this.value);"
                   onfocus="this.style.borderColor='#0078d4'; this.style.boxShadow='0 0 0 1px #0078d4';"
                   onblur="this.style.borderColor='#8a8886'; this.style.boxShadow='none';">
          </div>
        </div>
        
        <!-- Manifest GUID Field with Fluent UI styling -->
        <div style="margin-bottom: 12px;">
          <label class="ms-Label" style="display: block; margin-bottom: 4px; font-size: 14px; font-weight: 600; color: #323130; font-family: 'Segoe UI', sans-serif;">Manifest GUID:</label>
          <div class="ms-TextField">
            <input type="text" id="lib-guid-${index}" value="${this._escapeHtml(lib.manifestGuid)}" placeholder="e.g., 4588a19c-6c21-4f42-9bb6-9a7a3b62b1fa" 
                   class="ms-TextField-field"
                   style="width: calc(100% - 24px); padding: 8px 12px; border: 1px solid #8a8886; border-radius: 2px; font-size: 14px; font-family: 'Segoe UI Mono', monospace; background: #ffffff;" 
                   onchange="window.spfxExtensibilityActions.updateLibraryGuid(${index}, this.value);"
                   onfocus="this.style.borderColor='#0078d4'; this.style.boxShadow='0 0 0 1px #0078d4';"
                   onblur="this.style.borderColor='#8a8886'; this.style.boxShadow='none';">
          </div>
        </div>
        
        <div style="text-align: right;">
          <button type="button" class="ms-Button ms-Button--default" style="background: #d13438 !important; border: 1px solid #d13438 !important; color: white !important; padding: 6px 12px !important; font-size: 12px !important; border-radius: 2px !important; cursor: pointer !important; font-family: 'Segoe UI', sans-serif !important;" onclick="window.spfxExtensibilityActions.removeLibrary(${index});">
            <span class="ms-Button-label" style="font-weight: 400 !important;">Remove</span>
          </button>
        </div>
      </div>
    `).join('');
  }

  private _showSettingsPanel(): void {
    this._createSettingsPanel();
  }

  private _createExtensibilityPanel(): void {
    // Remove existing panel if any
    this._removePanel('extensibility-panel');

    // Get current libraries or create default ones
    const libraries = this._getExtensibilityLibraries();

    // Create panel overlay
    const overlay = document.createElement('div');
    overlay.id = 'extensibility-panel-overlay';
    overlay.className = 'ms-Overlay ms-Overlay--dark';
    overlay.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: 1000; background: rgba(0,0,0,0.4);';

    // Create panel
    const panel = document.createElement('div');
    panel.id = 'extensibility-panel';
    panel.className = 'ms-Panel ms-Panel--medium ms-Panel--right';
    panel.style.cssText = 'position: fixed; top: 0; right: 0; width: 600px; height: 100%; background: white; box-shadow: 0 0 20px rgba(0,0,0,0.3); z-index: 1001;';

    // Generate libraries HTML using the new method
    const librariesHTML = this._generateLibrariesHTML(libraries);

    // Panel content with proper scrolling structure
    panel.innerHTML = `
      <div class="ms-Panel-main" style="display: flex; flex-direction: column; height: 100%;">
        <div class="ms-Panel-commands" style="flex-shrink: 0;">
          <button type="button" class="ms-Panel-closeButton ms-PanelAction-close" style="border: none; background: transparent; padding: 8px; cursor: pointer;" onclick="document.getElementById('extensibility-panel-overlay').remove();">
            <span class="ms-Icon ms-Icon--Cancel" style="font-size: 16px; color: #605e5c;">✕</span>
          </button>
        </div>
        <div class="ms-Panel-contentInner" style="flex: 1; display: flex; flex-direction: column; overflow: hidden;">
          <div class="ms-Panel-header" style="flex-shrink: 0;">
            <p class="ms-Panel-headerText" style="font-size: 20px !important; font-weight: 600 !important; margin: 0 !important; padding: 20px 24px 0 24px !important; color: #323130 !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important;">Configure Extensibility Libraries</p>
          </div>
          <div class="ms-Panel-scrollableContent" style="flex: 1; overflow-y: auto; overflow-x: hidden; scrollbar-width: thin; scrollbar-color: #c8c6c4 #f3f2f1;">
            <div class="ms-Panel-content" style="padding: 20px 24px;">
              <p style="margin-bottom: 20px; color: #605e5c; font-size: 14px; font-family: 'Segoe UI', sans-serif;">
                Configure extensibility libraries that will be loaded with your search web part. These libraries can extend functionality with custom layouts, data sources, and search enhancements.
              </p>
              <div id="libraries-container" style="margin-bottom: 20px;">
                ${librariesHTML}
              </div>
              <div style="margin: 20px 0; text-align: center; padding: 16px 0; border-top: 1px solid #f3f2f1;">
                <button type="button" class="ms-Button ms-Button--primary" style="background: #0078d4 !important; border: 1px solid #0078d4 !important; color: white !important; padding: 8px 16px !important; font-size: 14px !important; border-radius: 2px !important; cursor: pointer !important; font-family: 'Segoe UI', sans-serif !important;" onclick="window.spfxExtensibilityActions.addNewLibrary();">
                  <span style="font-weight: 400 !important;">+ Add New Library</span>
                </button>
              </div>
            </div>
          </div>
          <div class="ms-Panel-footer" style="flex-shrink: 0; border-top: 1px solid #edebe9; padding: 16px 24px; background: #faf9f8;">
            <div style="display: flex; gap: 8px; justify-content: flex-end;">
              <button type="button" class="ms-Button ms-Button--primary" style="background: #0078d4 !important; border: 1px solid #0078d4 !important; color: white !important; padding: 8px 16px !important; font-size: 14px !important; border-radius: 2px !important; cursor: pointer !important; font-family: 'Segoe UI', sans-serif !important;" onclick="window.spfxExtensibilityActions.saveLibraries();">
                <span style="font-weight: 400 !important;">Apply Changes</span>
              </button>
              <button type="button" class="ms-Button" style="background: transparent !important; border: 1px solid #8a8886 !important; color: #323130 !important; padding: 8px 16px !important; font-size: 14px !important; border-radius: 2px !important; cursor: pointer !important; font-family: 'Segoe UI', sans-serif !important;" onclick="document.getElementById('extensibility-panel-overlay').remove();">
                <span style="font-weight: 400 !important;">Cancel</span>
              </button>
              <button type="button" class="ms-Button" style="background: transparent !important; border: 1px solid #8a8886 !important; color: #323130 !important; padding: 8px 16px !important; font-size: 14px !important; border-radius: 2px !important; cursor: pointer !important; font-family: 'Segoe UI', sans-serif !important;" onclick="window.spfxExtensibilityActions.resetToDefaults();">
                <span style="font-weight: 400 !important;">Reset to Defaults</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    `;

    // Setup global actions for panel interactions
    this._setupExtensibilityActions();

    // Add custom scrollbar styles for WebKit browsers
    const style = document.createElement('style');
    style.textContent = `
      #extensibility-panel .ms-Panel-scrollableContent::-webkit-scrollbar {
        width: 8px;
      }
      #extensibility-panel .ms-Panel-scrollableContent::-webkit-scrollbar-track {
        background: #f3f2f1;
        border-radius: 4px;
      }
      #extensibility-panel .ms-Panel-scrollableContent::-webkit-scrollbar-thumb {
        background: #c8c6c4;
        border-radius: 4px;
      }
      #extensibility-panel .ms-Panel-scrollableContent::-webkit-scrollbar-thumb:hover {
        background: #a19f9d;
      }
    `;
    document.head.appendChild(style);

    // Add to DOM
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    // Close panel when clicking overlay
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        overlay.remove();
      }
    });
  }

  private _createSettingsPanel(): void {
    // Remove existing panel if any
    this._removePanel('settings-panel');

    // Get current settings
    const currentSettings = { ...this.properties };
    delete (currentSettings as any).showExtensibilityPanel;
    delete (currentSettings as any).showSettingsPanel;

    // Create panel overlay
    const overlay = document.createElement('div');
    overlay.id = 'settings-panel-overlay';
    overlay.className = 'ms-Overlay ms-Overlay--dark';
    overlay.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: 1000; background: rgba(0,0,0,0.4);';

    // Create panel
    const panel = document.createElement('div');
    panel.id = 'settings-panel';
    panel.className = 'ms-Panel ms-Panel--large ms-Panel--right';
    panel.style.cssText = 'position: fixed; top: 0; right: 0; width: 644px; height: 100%; z-index: 1001; background: white; box-shadow: -6px 0 12px rgba(0,0,0,0.15);';

    // Panel content
    panel.innerHTML = `
      <div class="ms-Panel-main">
        <div class="ms-Panel-commands">
          <button type="button" class="ms-Panel-closeButton ms-PanelAction-close" onclick="document.getElementById('settings-panel-overlay').remove();" style="border: none; background: transparent; padding: 8px; cursor: pointer;">
            <span class="ms-Icon ms-Icon--Cancel" style="font-size: 16px; color: #605e5c;">✕</span>
          </button>
        </div>
        <div class="ms-Panel-contentInner">
          <div class="ms-Panel-header">
            <p class="ms-Panel-headerText" style="font-size: 20px !important; font-weight: 600 !important; margin: 0 !important; padding: 20px 24px 0 24px !important; color: #323130 !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important;">Edit Properties - JSON Configuration</p>
          </div>
          <div class="ms-Panel-scrollableContent">
            <div class="ms-Panel-content" style="padding: 20px 24px;">
              <p style="margin-bottom: 16px !important; color: #605e5c !important; font-size: 14px !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important;">
                Edit the web part configuration in JSON format. You can export current settings, modify them, and import back.
              </p>
              <textarea id="settings-json" 
                style="width: 100%; height: 400px; font-family: 'Segoe UI Mono', 'Courier New', monospace; font-size: 12px; border: 1px solid #edebe9; padding: 12px; border-radius: 2px; resize: vertical; background: #faf9f8;" 
                placeholder="JSON configuration will appear here...">${JSON.stringify(currentSettings, null, 2)}</textarea>
              <div style="margin-top: 20px; display: flex; gap: 8px; border-top: 1px solid #edebe9; padding-top: 16px;">
                <button type="button" class="ms-Button ms-Button--primary" style="background: #0078d4 !important; border: 1px solid #0078d4 !important; color: white !important; padding: 8px 16px !important; font-size: 14px !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important; border-radius: 2px !important; cursor: pointer !important;" onclick="window.spfxPanelActions.applySettings();">
                  <span class="ms-Button-label" style="font-weight: 400 !important;">Apply</span>
                </button>
                <button type="button" class="ms-Button" style="background: transparent !important; border: 1px solid #8a8886 !important; color: #323130 !important; padding: 8px 16px !important; font-size: 14px !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important; border-radius: 2px !important; cursor: pointer !important; margin-right: 8px !important;" onclick="document.getElementById('settings-panel-overlay').remove();">
                  <span class="ms-Button-label" style="font-weight: 400 !important;">Cancel</span>
                </button>
                <button type="button" class="ms-Button" style="background: transparent !important; border: 1px solid #8a8886 !important; color: #323130 !important; padding: 8px 16px !important; font-size: 14px !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important; border-radius: 2px !important; cursor: pointer !important; margin-right: 8px !important;" onclick="window.spfxPanelActions.exportSettings();">
                  <span class="ms-Button-label" style="font-weight: 400 !important;">Export</span>
                </button>
                <button type="button" class="ms-Button" style="background: transparent !important; border: 1px solid #8a8886 !important; color: #323130 !important; padding: 8px 16px !important; font-size: 14px !important; font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif !important; border-radius: 2px !important; cursor: pointer !important;" onclick="window.spfxPanelActions.importSettings();">
                  <span class="ms-Button-label" style="font-weight: 400 !important;">Import</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    // Add to DOM
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    // Setup global actions
    (window as any).spfxPanelActions = {
      applySettings: () => {
        try {
          const textarea = document.getElementById('settings-json') as HTMLTextAreaElement;
          const newSettings = JSON.parse(textarea.value);
          Object.assign(this.properties, newSettings);
          this.context.propertyPane.refresh();
          this.render();
          overlay.remove();
        } catch (error) {
          alert('Invalid JSON format. Please check your configuration.');
        }
      },
      exportSettings: () => {
        const textarea = document.getElementById('settings-json') as HTMLTextAreaElement;
        const blob = new Blob([textarea.value], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spfx-banner-search-settings.json';
        a.click();
        URL.revokeObjectURL(url);
      },
      importSettings: () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json';
        input.onchange = (e) => {
          const file = (e.target as HTMLInputElement).files?.[0];
          if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
              const textarea = document.getElementById('settings-json') as HTMLTextAreaElement;
              textarea.value = e.target?.result as string;
            };
            reader.readAsText(file);
          }
        };
        input.click();
      }
    };

    // Close on overlay click
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        overlay.remove();
      }
    });
  }

  private _removePanel(panelId: string): void {
    const existingOverlay = document.getElementById(`${panelId}-overlay`);
    if (existingOverlay) {
      existingOverlay.remove();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure the hero banner appearance and search box"
          },
          groups: [
            {
              groupName: "Banner & Visual Settings",
              groupFields: [
                PropertyFieldColorPicker('gradientStartColor', {
                  label: 'Gradient Start Color',
                  selectedColor: this.properties.gradientStartColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'gradientStartColorFieldId'
                }),
                PropertyFieldColorPicker('gradientEndColor', {
                  label: 'Gradient End Color',
                  selectedColor: this.properties.gradientEndColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'gradientEndColorFieldId'
                }),
                PropertyPaneToggle('showCircleAnimation', {
                  label: 'Show Circle Animation',
                  checked: this.properties.showCircleAnimation
                }),
                PropertyPaneSlider('minHeight', {
                  label: 'Banner Minimum Height (px)',
                  min: 350,
                  max: 800,
                  step: 50,
                  showValue: true,
                  value: this.properties.minHeight
                }),
                PropertyPaneSlider('titleFontSize', {
                  label: 'Title Font Size (px)',
                  min: 24,
                  max: 72,
                  step: 4,
                  showValue: true,
                  value: this.properties.titleFontSize
                }),
                PropertyFieldColorPicker('bannerTitleColor', {
                  label: 'Banner Title Color',
                  selectedColor: this.properties.bannerTitleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Font',
                  key: 'bannerTitleColorFieldId'
                }),
                PropertyPaneTextField('bannerTitle', {
                  label: 'Banner Title',
                  description: 'Title text. Use {email}, {firstname}, {lastname}, {displayname} for dynamic user properties',
                  value: this.properties.bannerTitle,
                  placeholder: 'Find What You Need'
                }),
                PropertyPaneTextField('searchBoxPlaceholder', {
                  label: 'Search Box Placeholder Text',
                  value: this.properties.searchBoxPlaceholder
                }),
                PropertyPaneSlider('searchBoxBorderRadius', {
                  label: 'Search Box Border Radius (px)',
                  min: 0,
                  max: 15,
                  step: 1,
                  showValue: true,
                  value: this.properties.searchBoxBorderRadius || 4
                }),
                PropertyPaneSlider('searchBoxHeight', {
                  label: 'Search Box Height (px)',
                  min: 30,
                  max: 50,
                  step: 2,
                  showValue: true,
                  value: this.properties.searchBoxHeight || 32
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Configure search behavior and connections"
          },
          groups: [
            {
              groupName: "Search Configuration",
              groupFields: [
                PropertyPaneTextField('queryTemplate', {
                  label: 'Query Template',
                  description: 'Default query template to use (e.g., * for all results)',
                  value: this.properties.queryTemplate
                }),
                PropertyPaneTextField('resultsWebPartId', {
                  label: 'Results Web Part ID (Optional)',
                  description: 'GUID of the Search Results web part to connect to',
                  value: this.properties.resultsWebPartId
                }),
                PropertyPaneToggle('enableSuggestions', {
                  label: 'Enable Search Suggestions',
                  checked: this.properties.enableSuggestions
                }),
                PropertyPaneSlider('suggestionsLimit', {
                  label: 'Number of suggestions to show per group',
                  min: 1,
                  max: 20,
                  step: 1,
                  showValue: true,
                  value: this.properties.suggestionsLimit || 10,
                  disabled: this.properties.enableSuggestions === false
                }),
                PropertyPaneDropdown('openingBehavior', {
                  label: 'Opening behavior',
                  options: [
                    { key: 'current-tab', text: 'Use the current tab' },
                    { key: 'new-tab', text: 'Open in the new tab' }
                  ],
                  selectedKey: this.properties.openingBehavior || 'new-tab',
                  disabled: this.properties.enableSuggestions === false
                })
              ]
            },
            {
              groupName: "Search Behavior",
              groupFields: [
                PropertyPaneToggle('redirectToSearchPage', {
                  label: 'Redirect to Search Page',
                  checked: this.properties.redirectToSearchPage || false
                }),
                PropertyPaneTextField('searchPageUrl', {
                  label: 'Search Page URL',
                  description: 'URL to redirect to when search is submitted (e.g., /sites/search/pages/results.aspx)',
                  value: this.properties.searchPageUrl || '',
                  disabled: !this.properties.redirectToSearchPage
                }),
                PropertyPaneDropdown('searchMethod', {
                  label: 'Method',
                  options: [
                    { key: 'query-string', text: 'Query string parameter' },
                    { key: 'url-fragment', text: 'URL fragment' }
                  ],
                  selectedKey: this.properties.searchMethod || 'query-string',
                  disabled: !this.properties.redirectToSearchPage
                }),
                PropertyPaneTextField('searchParameterName', {
                  label: 'Parameter name',
                  description: 'URL parameter name for the search query (e.g., q, search, query)',
                  value: this.properties.searchParameterName || 'q',
                  placeholder: 'q',
                  disabled: !this.properties.redirectToSearchPage || this.properties.searchMethod === 'url-fragment'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Configure query suggestions and search experience"
          },
          groups: [
            {
              groupName: "Query Suggestions",
              groupFields: [
                PropertyPaneToggle('enableQuerySuggestions', {
                  label: 'Enable Query Suggestions',
                  checked: this.properties.enableQuerySuggestions !== false
                }),
                PropertyPaneDropdown('suggestionsProvider', {
                  label: 'Suggestions Provider',
                  options: [
                    { key: 'static', text: 'SharePoint Static Suggestions' },
                    { key: 'search', text: 'SharePoint Search Suggestions' },
                    { key: 'custom', text: 'Custom Provider' }
                  ],
                  selectedKey: this.properties.suggestionsProvider || 'static',
                  disabled: this.properties.enableQuerySuggestions === false
                }),
                PropertyPaneTextField('staticSuggestions', {
                  label: 'Static Suggestions',
                  description: 'Enter suggestions separated by commas (e.g., "policy, procedure, guidelines")',
                  value: this.properties.staticSuggestions || '',
                  multiline: true,
                  rows: 5,
                  disabled: this.properties.enableQuerySuggestions === false || this.properties.suggestionsProvider !== 'static'
                }),
                PropertyPaneToggle('enableZeroTermSuggestions', {
                  label: 'Show suggestions when search box is empty',
                  checked: this.properties.enableZeroTermSuggestions !== false,
                  disabled: this.properties.enableQuerySuggestions === false
                }),
                PropertyPaneTextField('zeroTermSuggestions', {
                  label: 'Zero-term Suggestions',
                  description: 'Suggestions to show when search box is empty, separated by commas',
                  value: this.properties.zeroTermSuggestions || '',
                  multiline: true,
                  rows: 3,
                  disabled: this.properties.enableQuerySuggestions === false || this.properties.enableZeroTermSuggestions === false
                })
              ]
            },
            {
              groupName: "Custom Search Suggestions",
              groupFields: [
                PropertyPaneTextField('hubSiteId', {
                  label: 'Hub Site ID',
                  description: 'Enter the Hub Site ID for custom search suggestions',
                  value: this.properties.hubSiteId || '',
                  placeholder: 'e.g., 12345678-1234-1234-1234-123456789012',
                  disabled: this.properties.enableQuerySuggestions === false || this.properties.suggestionsProvider !== 'custom'
                }),
                PropertyPaneTextField('imageRelativeUrl', {
                  label: 'Image Relative URL',
                  description: 'Relative URL for suggestion images (e.g., /sites/hub/images/suggestions.png)',
                  value: this.properties.imageRelativeUrl || '',
                  placeholder: '/sites/hub/images/suggestions.png',
                  disabled: this.properties.enableQuerySuggestions === false || this.properties.suggestionsProvider !== 'custom'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Configure dynamic data connections and page environment settings"
          },
          groups: [
            {
              groupName: "Dynamic Data Source",
              groupFields: [
                PropertyPaneToggle('useDynamicDataSource', {
                  label: 'Use dynamic data source',
                  checked: this.properties.useDynamicDataSource !== false
                }),
                PropertyPaneDropdown('dynamicDataSourceId', {
                  label: 'Connect to source',
                  options: [
                    { key: 'pageEnvironment', text: 'Page Environment' }
                  ],
                  selectedKey: this.properties.dynamicDataSourceId || 'pageEnvironment',
                  disabled: !this.properties.useDynamicDataSource
                }),
                PropertyPaneDropdown('pageEnvironmentProperty', {
                  label: 'Page environment\'s properties',
                  options: [
                    { key: 'siteProperties', text: 'Site Properties' },
                    { key: 'currentUser', text: 'Current User Information' },
                    { key: 'queryString', text: 'Query String' },
                    { key: 'search', text: 'Search' }
                  ],
                  selectedKey: this.properties.pageEnvironmentProperty || '',
                  disabled: !this.properties.useDynamicDataSource || this.properties.dynamicDataSourceId !== 'pageEnvironment'
                }),
                // Site Properties dropdown - only show if site properties is selected
                ...(this.properties.pageEnvironmentProperty === 'siteProperties' && this.properties.useDynamicDataSource ? [
                  PropertyPaneDropdown('siteProperty', {
                    label: 'Site property',
                    options: [
                      { key: 'siteUrl', text: 'Site URL' },
                      { key: 'siteCollectionUrl', text: 'Site Collection URL' },
                      { key: 'siteTitle', text: 'Site Title' },
                      { key: 'siteId', text: 'Site ID' },
                      { key: 'webId', text: 'Web ID' },
                      { key: 'hubSiteId', text: 'Hub Site ID' }
                    ],
                    selectedKey: this.properties.siteProperty || '',
                    disabled: !this.properties.useDynamicDataSource
                  })
                ] : []),
                // Current User dropdown - only show if current user is selected
                ...(this.properties.pageEnvironmentProperty === 'currentUser' && this.properties.useDynamicDataSource ? [
                  PropertyPaneDropdown('userProperty', {
                    label: 'User property',
                    options: [
                      { key: 'loginName', text: 'Login Name' },
                      { key: 'displayName', text: 'Display Name' },
                      { key: 'email', text: 'Email' },
                      { key: 'userId', text: 'User ID' },
                      { key: 'department', text: 'Department' },
                      { key: 'jobTitle', text: 'Job Title' }
                    ],
                    selectedKey: this.properties.userProperty || '',
                    disabled: !this.properties.useDynamicDataSource
                  })
                ] : []),
                // Query String dropdown - only show if query string is selected
                ...(this.properties.pageEnvironmentProperty === 'queryString' && this.properties.useDynamicDataSource ? [
                  PropertyPaneTextField('queryStringProperty', {
                    label: 'Query string parameter name',
                    description: 'Enter the name of the query string parameter to use (e.g., q, search, query)',
                    value: this.properties.queryStringProperty || '',
                    placeholder: 'q',
                    disabled: !this.properties.useDynamicDataSource
                  })
                ] : []),
                // Search dropdown - only show if search is selected
                ...(this.properties.pageEnvironmentProperty === 'search' && this.properties.useDynamicDataSource ? [
                  PropertyPaneDropdown('searchProperty', {
                    label: 'Search property',
                    options: [
                      { key: 'searchQuery', text: 'Search Query' },
                      { key: 'searchResults', text: 'Search Results' },
                      { key: 'selectedFilters', text: 'Selected Filters' },
                      { key: 'searchVertical', text: 'Search Vertical' }
                    ],
                    selectedKey: this.properties.searchProperty || '',
                    disabled: !this.properties.useDynamicDataSource
                  })
                ] : [])
              ]
            }
          ]
        },
        {
          header: {
            description: "About this web part, extensibility resources, and advanced settings"
          },
          groups: [
            {
              groupName: "About",
              groupFields: [
                PropertyPaneTextField('', {
                  label: 'Author',
                  value: 'Monarch360',
                  disabled: true,
                  description: 'Visit https://www.monarch360.com.au/'
                }),
                PropertyPaneTextField('', {
                  label: 'Developer',
                  value: 'Shahanur Sharif',
                  disabled: true
                }),
                PropertyPaneTextField('', {
                  label: 'Version',
                  value: packageInfo.version || '1.0.0',
                  disabled: true
                }),
                PropertyPaneTextField('', {
                  label: 'Web Part Instance ID',
                  value: this.context.instanceId,
                  disabled: true,
                  description: 'Unique identifier for this web part instance'
                })
              ]
            },
            {
              groupName: "Resources",
              groupFields: [
                PropertyPaneButton('showExtensibilityPanel', {
                  text: 'Extensibility libraries to load',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    this._showExtensibilityPanel();
                  }
                })
              ]
            },
            {
              groupName: "Export/Import Settings",
              groupFields: [
                PropertyPaneButton('showSettingsPanel', {
                  text: 'Edit Properties',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => {
                    this._showSettingsPanel();
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
