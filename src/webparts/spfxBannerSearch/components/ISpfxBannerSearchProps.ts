import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISemanticColors } from '@fluentui/react/lib/Styling';

export interface ISpfxBannerSearchProps {
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
  
  // About and settings configuration - panels now handled natively in web part
  
  // Extensibility libraries configuration

  
  // SPFx context and theme
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  semanticColors: Partial<ISemanticColors>;
  context: WebPartContext;
  onSearchQuery: (queryText: string) => void;

  /**
   * Optional: Override the site URL for search suggestions (defaults to context.pageContext.web.absoluteUrl)
   */
  searchSiteUrl?: string;

  /**
   * Optional: Enable debug logging for type-ahead suggestions
   */
  debugSuggestions?: boolean;
}
