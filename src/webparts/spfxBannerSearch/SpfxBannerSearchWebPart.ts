import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISemanticColors } from '@fluentui/react/lib/Styling';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

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
        }
      ]
    };
  }
}
