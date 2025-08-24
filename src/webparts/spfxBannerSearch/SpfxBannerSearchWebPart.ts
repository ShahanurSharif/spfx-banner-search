import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
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
  bannerTitle: string;
  searchBoxPlaceholder: string;
  
  // Search behavior configuration
  queryTemplate: string;
  resultsWebPartId: string;
  enableSuggestions: boolean;
  
  // Redirect configuration
  redirectToSearchPage: boolean;
  searchPageUrl: string;
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
        minHeight: this.properties.minHeight || 500,
        bannerTitle: this._processDynamicTitle(this.properties.bannerTitle || 'Find What You Need'),
        searchBoxPlaceholder: this.properties.searchBoxPlaceholder || 'Search everything...',
        queryTemplate: this.properties.queryTemplate || '*',
        enableSuggestions: this.properties.enableSuggestions !== false,
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
      const searchUrl = `${this.properties.searchPageUrl}?q=${encodeURIComponent(queryText)}`;
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
                  min: 500,
                  max: 800,
                  step: 50,
                  showValue: true,
                  value: this.properties.minHeight
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
